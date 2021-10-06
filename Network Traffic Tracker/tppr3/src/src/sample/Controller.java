package src.sample;

import com.google.gson.Gson;
import com.google.gson.JsonObject;
import com.rabbitmq.client.*;
import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.fxml.Initializable;
import javafx.scene.control.Button;
import javafx.scene.control.PasswordField;
import javafx.scene.control.ProgressBar;
import javafx.scene.control.TextField;
import sun.rmi.transport.Channel;
import sun.rmi.transport.Connection;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URL;
import java.util.ResourceBundle;
import java.util.concurrent.TimeoutException;

public class Controller implements Consumer, Initializable {


    public static String idCSV = "";
    public static String tsCSV = "";
    public static String valueCSV = "";
    public static int count;
    public static boolean stop = false;
    private final long MILLION = 1000000;
    private final int MILLIS_IN_ON_SECONDS = 1000;

    private long startTimeInMilli;
    private long endTimeInMilli;
    
    //rabbit vars
    public static com.rabbitmq.client.Channel dataChannel;
    public static String exchange = "amq.topic";
    public static String dataQueue;
    public static com.rabbitmq.client.Connection connection = null;


    public static boolean stopped;
    private String machineId;
    private int durationInMilli;
    private String mqServerUrl;
    private String username;
    private String password;


    @FXML
    public  ProgressBar progressStatus;
    @FXML
    private Button start;
    @FXML
    private TextField machIdText;
    @FXML
    public TextField serverAddressText;
    @FXML
    public TextField usernameText;
    @FXML
    public PasswordField passwordText;
    @FXML
    private TextField durationText;
    @FXML
    private TextField errorText;
    private Object Exception;


    @FXML
    void startCapture(ActionEvent event) {
        progressStatus.setVisible(true);
        errorText.setVisible(false);
        progressStatus.setProgress(0.05);
        ConnectionFactory factory = new ConnectionFactory();

        try {
            this.setVars(this, machIdText.getText(), durationText.getText(), serverAddressText.getText(), usernameText.getText(), passwordText.getText(), factory);
        } catch (Exception e) {
            errorText.setVisible(true);
            errorText.setText("Please check Server Settings and write permissions");


        }

    }


    public void setVars(Controller controller, String machineId, String duration, String mqServerUrl, String username, String password, ConnectionFactory factory) throws Exception {


        controller.machineId = machineId;
        controller.durationInMilli = Integer.parseInt(duration);//in secoinds
        controller.mqServerUrl = mqServerUrl;
        controller.username = username;
        controller.password = password;

        Controller.startTapping(this, factory);

    }


    public static void startTapping(Controller controller, ConnectionFactory factory) throws Exception {


        factory.setHost(controller.mqServerUrl);
        factory.setUsername(controller.username);
        factory.setPassword(controller.password);
        factory.setPort(5672);


            Controller.connection = factory.newConnection();


        controller.startTimeInMilli = System.currentTimeMillis();
        controller.endTimeInMilli = controller.startTimeInMilli + controller.durationInMilli * controller.MILLIS_IN_ON_SECONDS;
        Controller.stopped = false;


            Controller.dataChannel = Controller.connection.createChannel();
            Controller.dataChannel.exchangeDeclare(Controller.exchange, "topic", true);
            Controller.dataQueue = Controller.dataChannel.queueDeclare("ExcelDataTapperQueue1", true, false, false, null).getQueue();
            Controller.dataChannel.queueBind(Controller.dataQueue, Controller.exchange, "comau.smart7." + controller.machineId + ".data.plc");

            //startConsumption

            Controller.dataChannel.basicConsume(Controller.dataQueue, true, controller);

    }


    public void handleDelivery(String consumerTag, Envelope envelope, AMQP.BasicProperties props, byte[] msg) throws IOException {



        if (Controller.stopped) {
           //ed Message");
            return;
        } else {
            progressStatus.setProgress(0.15);
            //System.out.println("Got Message ");
        }

        String message = new String(msg);

        System.out.println("Received Message" + message);

        if (System.currentTimeMillis() > this.endTimeInMilli) {
            Controller.stopped = true;


            this.stopListening(consumerTag);
            try {
                this.writeToExcel(csvToStringList(Controller.idCSV), csvToStringList(Controller.tsCSV), csvToStringList(Controller.valueCSV));
            } catch (Throwable throwable) {
                throwable.printStackTrace();
            }


        }


        if (!Controller.stopped) {

            Controller.idCSV = Controller.idCSV + Controller.removeQuotes(Controller.getStringList(message, "id")) + ",";
            Controller.tsCSV = Controller.tsCSV + Controller.removeQuotes(Controller.getStringList(message, "ts")) + ",";
            Controller.valueCSV = Controller.valueCSV + Controller.removeQuotes(Controller.getStringList(message, "v")) + ",";
        } 
    }



    public static String getStringList(String dataElement, String elementName) {

        Gson gson = new Gson();
        return gson.fromJson(dataElement, JsonObject.class).get(elementName).toString().replace("[", "").replace("]", "");

    }

    private void writeToExcel(String[] idArr, String[] tsArr, String[] valueArr) throws Throwable {

        System.out.println("Writing To Excel");

        progressStatus.setProgress( 0.35);
       
        String[] columns = {"NAME", "TS", "VALUE"};
        XSSFWorkbook workbook = null;

            workbook = new XSSFWorkbook();


		/* CreationHelper helps us create instances of various things like DataFormat,
           Hyperlink, RichTextString etc, in a format (HSSF, XSSF) independent way */
        //CreationHelper createHelper = workbook.getCreationHelper();

        // Create a Sheet
        Sheet sheet = workbook.createSheet("PLC_DATA");

        // Create a Font for styling header cells
        Font headerFont = workbook.createFont();
        //headerFont.setBoldweight((short) 1);
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 12);
        headerFont.setColor(IndexedColors.GREY_80_PERCENT.getIndex());

        // Create a CellStyle with the font
        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);

        // Create a Row
        Row headerRow = sheet.createRow(0);

        // Create cells
        for (int i = 0; i < columns.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(columns[i]);
            cell.setCellStyle(headerCellStyle);

        }

        //excluding header
        int rowNum = 1;

        DataFormat format = workbook.createDataFormat();

        // Create a CellStyle for ts
        CellStyle tsStyle = workbook.createCellStyle();
        tsStyle.setDataFormat(format.getFormat("@"));


        // Create a CellStyle for values
        CellStyle numberStyle = workbook.createCellStyle();

        numberStyle.setDataFormat(format.getFormat("@"));


        for (int i = 0; i < idArr.length; i++) {
            Row row = sheet.createRow(rowNum++);

            row.createCell(0).setCellValue(idArr[i]);

            Cell tsCell = row.createCell(1);
            tsCell.setCellValue(tsArr[i]);
            tsCell.setCellStyle(tsStyle);


            Cell valueCell = row.createCell(1);
            valueCell.setCellValue(tsArr[i]);
            valueCell.setCellStyle(numberStyle);


            row.createCell(2).setCellValue(valueArr[i]);
            System.out.println("id:" +idArr[i] + "  ts:" + tsArr[i]+ "  value:"+ valueArr[i]);

            if (i > 1000000) {//excel row limt
                i = idArr.length;
                throw (Throwable) Exception;
               // System.out.println("Rows Exceeded 1/,000/,000, data will be truncated in excel, due to row size limitation!");


            }

        }
        progressStatus.setProgress(0.75);

       
        // Resize all columns to fit the content size
        for (int i = 0; i < columns.length; i++) {
            sheet.autoSizeColumn(i);
        }
        progressStatus.setProgress(0.85);
      
        // Write the output to a file
        long epoch = System.currentTimeMillis();
        String FileName = String.valueOf(epoch);
        FileOutputStream fileOut = new FileOutputStream("C:\\Users\\Admin\\Desktop\\Project 19-20\\Report" + FileName + ".xlsx");
        workbook.write(fileOut);
        fileOut.close();
        progressStatus.setProgress(1);
        // Closing the workbook
        workbook.close();

    }

    public static String removeQuotes(String in) {
        return in.replaceAll("\"", "");

    }


    private static String[] csvToStringList(String idCSV) {

        return (idCSV.split(","));

    }

    public void handleConsumeOk(String consumerTag) {
        // TODO Auto-generated method stub

    }


    public void handleCancelOk(String consumerTag) {
        // TODO Auto-generated method stub

    }


    public void handleCancel(String consumerTag) throws IOException {
        // TODO Auto-generated method stub

    }


    public void handleShutdownSignal(String consumerTag, ShutdownSignalException sig) {


    }


    public void handleRecoverOk(String consumerTag) {
        // TODO Auto-generated method stub

    }


    public void stopListening(String consumerTag) {

        try {

            Controller.dataChannel.close();
            Controller.connection.close();



            // connection.close();
        } catch (IOException e) {

            e.printStackTrace();
        } catch (TimeoutException e) {

            e.printStackTrace();
        }

       

    }



    @Override
    public void initialize(URL location, ResourceBundle resources) {
        start.defaultButtonProperty().bind(start.focusedProperty());
        errorText.setVisible(false);
        progressStatus.setVisible(false);
        progressStatus.setProgress(0.01);
    }


	public void initialize1(URL arg0, ResourceBundle arg1) {
		// TODO Auto-generated method stub
		
	}
}
