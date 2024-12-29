import javafx.application.Application;
import javafx.geometry.HPos;
import javafx.geometry.Insets;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.TextField;
import javafx.scene.layout.BorderPane;
import javafx.scene.layout.ColumnConstraints;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.Priority;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.StringUtil;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Optional;

public class Main extends Application {
    private static final DateFormat dateTimeFormatter = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");

    private static File JIN_FILE;

    private static File MAO_FILE;

    public static void main(String[] args) {
        launch(args);
    }

    public static File onSaveClick() throws IOException, InvalidFormatException, ParseException {
        DataFormatter dataFormatter = new DataFormatter();
        Map<String, Model> id2Jin = new LinkedHashMap<>();
        XSSFWorkbook wb = new XSSFWorkbook(JIN_FILE);
        Sheet sheet = wb.getSheetAt(0);
        int rowIndex = 5;
        while (true) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) {
                break;
            }
            Cell cell = row.getCell(0);
            if (cell == null) {
                break;
            }
            String dataTimeStr = dataFormatter.formatCellValue(cell);
            if (dataTimeStr == null || StringUtil.isBlank(dataTimeStr)) {
                break;
            }

            String weightStr = dataFormatter.formatCellValue(row.getCell(3));

            String id = dataFormatter.formatCellValue(row.getCell(8));

            Model model = new Model();
            model.dateTime = dateTimeFormatter.parse(dataTimeStr);
            model.weightStr = weightStr;
            model.id = id;

            Model pre = id2Jin.get(id);
            if (pre == null || model.dateTime.after(pre.dateTime)) {
                id2Jin.put(id, model);
            }

            rowIndex++;
        }

        Map<String, Model> id2Mao = new LinkedHashMap<>();
        wb = new XSSFWorkbook(MAO_FILE);
        sheet = wb.getSheetAt(0);
        rowIndex = 5;
        while (true) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) {
                break;
            }
            Cell cell = row.getCell(0);
            if (cell == null) {
                break;
            }
            String dataTimeStr = dataFormatter.formatCellValue(cell);
            if (dataTimeStr == null || StringUtil.isBlank(dataTimeStr)) {
                break;
            }

            String weightStr = dataFormatter.formatCellValue(row.getCell(3));

            String id = dataFormatter.formatCellValue(row.getCell(8));

            Model model = new Model();
            model.dateTime = dateTimeFormatter.parse(dataTimeStr);
            model.weightStr = weightStr;
            model.id = id;

            Model pre = id2Mao.get(id);
            if (pre == null || model.dateTime.after(pre.dateTime)) {
                id2Mao.put(id, model);
            }

            rowIndex++;
        }


        wb = new XSSFWorkbook();
        sheet = wb.createSheet();

        XSSFCellStyle style = wb.createCellStyle();
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());

        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);
        cell.setCellValue("批号");
        cell.setCellStyle(style);

        cell = row.createCell(1);
        cell.setCellValue("客户名称");
        cell.setCellStyle(style);

        cell = row.createCell(2);
        cell.setCellValue("净重");
        cell.setCellStyle(style);

        cell = row.createCell(3);
        cell.setCellValue("毛重");
        cell.setCellStyle(style);

        cell = row.createCell(4);
        cell.setCellValue("数据累计");
        cell.setCellStyle(style);

        cell = row.createCell(5);
        cell.setCellValue("产品名称");
        cell.setCellStyle(style);

        cell = row.createCell(6);
        cell.setCellValue("品号");
        cell.setCellStyle(style);

        cell = row.createCell(7);
        cell.setCellValue("规格");
        cell.setCellStyle(style);

        cell = row.createCell(8);
        cell.setCellValue("质保期");
        cell.setCellStyle(style);

        cell = row.createCell(9);
        cell.setCellValue("生产日期");
        cell.setCellStyle(style);

        cell = row.createCell(10);
        cell.setCellValue("二维码编号");
        cell.setCellStyle(style);

        cell = row.createCell(11);
        cell.setCellValue("二维码材质");
        cell.setCellStyle(style);

        cell = row.createCell(12);
        cell.setCellValue("二维码净重");
        cell.setCellStyle(style);

        cell = row.createCell(13);
        cell.setCellValue("二维码毛重");
        cell.setCellStyle(style);

        cell = row.createCell(14);
        cell.setCellValue("二维码公司名称");
        cell.setCellStyle(style);

        cell = row.createCell(15);
        cell.setCellValue("二维码资料库");
        cell.setCellStyle(style);

        cell = row.createCell(16);
        cell.setCellValue("条形码信息");
        cell.setCellStyle(style);

        rowIndex = 1;
        for (Model model : id2Jin.values()) {
            if (!id2Mao.containsKey(model.id)) {
                continue;
            }

            row = sheet.createRow(rowIndex++);
            cell = row.createCell(0);
            cell.setCellValue(model.id.replace("PH ", ""));
            cell.setCellStyle(style);

            cell = row.createCell(1);
            cell.setCellStyle(style);

            cell = row.createCell(2);
            cell.setCellValue(model.weightStr);
            cell.setCellStyle(style);

            cell = row.createCell(3);
            cell.setCellValue(id2Mao.getOrDefault(model.id, new Model()).weightStr);
            cell.setCellStyle(style);

            cell = row.createCell(4);
            String shuJuLeiJi = model.weightStr.replace(" g", "");
            cell.setCellStyle(style);
            cell.setCellValue(shuJuLeiJi);

            cell = row.createCell(5);
            cell.setCellStyle(style);
            cell = row.createCell(6);
            cell.setCellStyle(style);
            cell = row.createCell(7);
            cell.setCellStyle(style);
            cell = row.createCell(8);
            cell.setCellStyle(style);

            cell = row.createCell(9);
            cell.setCellValue("20" + model.id.substring(6, 8) + "/" + model.id.substring(8, 10) + "/" + model.id.substring(10, 12));
            cell.setCellStyle(style);

            cell = row.createCell(10);
            cell.setCellValue("YDAu" + model.id.substring(5));
            cell.setCellStyle(style);

            cell = row.createCell(11);
            cell.setCellValue("AU,");
            cell.setCellStyle(style);

            cell = row.createCell(12);
            cell.setCellValue(model.weightStr.replace(" g", ","));
            cell.setCellStyle(style);

            cell = row.createCell(13);

            cell.setCellValue(Optional.ofNullable(id2Mao.getOrDefault(model.id, new Model()).weightStr).orElse(" g").replace(" g", ","));
            cell.setCellStyle(style);

            cell = row.createCell(14);
            cell.setCellValue("福兆二维码信息");
            cell.setCellStyle(style);

            cell = row.createCell(15);
            cell.setCellValue(row.getCell(10).getStringCellValue() + "-4N,AU," + row.getCell(12).getStringCellValue() + row.getCell(13).getStringCellValue());
            cell.setCellStyle(style);

            cell = row.createCell(16);
            cell.setCellValue(row.getCell(4).getStringCellValue());
            cell.setCellStyle(style);
        }

        for (int i = 0; i < 17; i++) {
            sheet.autoSizeColumn(i);
            sheet.setColumnWidth(i, sheet.getColumnWidth(i) * 17 / 10);
        }

        File output = new File("D:\\" + System.currentTimeMillis() + ".xlsx");
        try (FileOutputStream out = new FileOutputStream(output)) {
            wb.write(out);
        }
        return output;
    }

    @Override
    public void start(Stage primaryStage) {
        BorderPane root = new BorderPane();
        Scene scene = new Scene(root, 700, 90, javafx.scene.paint.Color.WHITE);
        GridPane gridpane = new GridPane();
        gridpane.setPadding(new Insets(5));
        gridpane.setHgap(5);
        gridpane.setVgap(5);
        ColumnConstraints column1 = new ColumnConstraints(150);
        ColumnConstraints column2 = new ColumnConstraints(50, 500, 500);
        column2.setHgrow(Priority.ALWAYS);
        gridpane.getColumnConstraints().addAll(column1, column2);

        Button jinZhongButton = new Button("选择净重表格文件");


        Button maoZhongButton = new Button("选择毛重表格文件");

        TextField jinZhongText = new TextField();
        jinZhongText.setEditable(false);
        TextField maoZhongText = new TextField();
        maoZhongText.setEditable(false);
        TextField saveText = new TextField();
        maoZhongText.setEditable(false);

        Button saveButton = new Button("          生成          ");

        jinZhongButton.setOnAction(event -> {
            FileChooser fileChooser = new FileChooser();
            FileChooser.ExtensionFilter extFilter = new FileChooser.ExtensionFilter("excel", "*.xlsx");
            fileChooser.getExtensionFilters().add(extFilter);
            JIN_FILE = fileChooser.showOpenDialog(primaryStage);
            jinZhongText.setText(JIN_FILE.getAbsolutePath());
        });

        maoZhongButton.setOnAction(event -> {
            FileChooser fileChooser = new FileChooser();
            FileChooser.ExtensionFilter extFilter = new FileChooser.ExtensionFilter("excel", "*.xlsx");
            fileChooser.getExtensionFilters().add(extFilter);
            MAO_FILE = fileChooser.showOpenDialog(primaryStage);
            maoZhongText.setText(MAO_FILE.getAbsolutePath());
        });

        saveButton.setOnAction(event -> {
            saveText.clear();
            try {
                File output = onSaveClick();
                saveText.setText("生成文件成功：" + output.getAbsolutePath());
            } catch (Exception e) {
                e.printStackTrace();
                saveText.setText("生成文件失败");
            }
        });

        GridPane.setHalignment(jinZhongButton, HPos.RIGHT);
        gridpane.add(jinZhongButton, 0, 0);

        GridPane.setHalignment(jinZhongText, HPos.LEFT);
        gridpane.add(jinZhongText, 1, 0);

        GridPane.setHalignment(maoZhongButton, HPos.RIGHT);
        gridpane.add(maoZhongButton, 0, 1);

        GridPane.setHalignment(maoZhongText, HPos.LEFT);
        gridpane.add(maoZhongText, 1, 1);

        GridPane.setHalignment(saveButton, HPos.RIGHT);
        gridpane.add(saveButton, 0, 2);

        GridPane.setHalignment(saveText, HPos.LEFT);
        gridpane.add(saveText, 1, 2);

        root.setCenter(gridpane);
        primaryStage.setScene(scene);
        primaryStage.show();

    }

    private static class Model {
        Date dateTime;

        String weightStr;

        String id;

        @Override
        public String toString() {
            return "Model{" +
                    "dateTimeStr='" + dateTime + '\'' +
                    ", weightStr='" + weightStr + '\'' +
                    ", id='" + id + '\'' +
                    '}';
        }
    }
}


