import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.telegram.telegrambots.ApiContextInitializer;
import org.telegram.telegrambots.bots.DefaultBotOptions;
import org.telegram.telegrambots.bots.TelegramLongPollingBot;
import org.telegram.telegrambots.meta.ApiContext;
import org.telegram.telegrambots.meta.TelegramBotsApi;
import org.telegram.telegrambots.meta.api.methods.send.SendMessage;
import org.telegram.telegrambots.meta.api.objects.Message;
import org.telegram.telegrambots.meta.api.objects.Update;
import org.telegram.telegrambots.meta.exceptions.TelegramApiException;

import java.io.FileInputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;

public class MyBot extends TelegramLongPollingBot {

    static XSSFWorkbook myExcelBook;


    public static void main(String[] args) throws Exception{

        ApiContextInitializer.init();
        TelegramBotsApi botsApi = new TelegramBotsApi();
        try {
            String file = "your/file/directory";
            myExcelBook = new XSSFWorkbook(new FileInputStream(file));
            DefaultBotOptions botOptions = ApiContext.getInstance(DefaultBotOptions.class);
            String PROXY_HOST = "78.46.200.216";
            botOptions.setProxyHost(PROXY_HOST);
            int PROXY_PORT = 34798;
            botOptions.setProxyPort(PROXY_PORT);
            botOptions.setProxyType(DefaultBotOptions.ProxyType.SOCKS5);
            botsApi.registerBot(new MyBot(botOptions));
        } catch (TelegramApiException e) {
            e.printStackTrace();
        }
    }

    public MyBot(DefaultBotOptions options) {
        super(options);
    }

    public String getBotUsername() {
        return "TaxcomBot";
    }

    public void onUpdateReceived(Update e) {
        Message msg = e.getMessage();
        String txt = msg.getText();
        if (txt.equals("/start")) {
            sendMsg(msg, "Привет работник Taxcom! Я помогу узнать тебе свое расписание," +
                    " просто напиши сообщение в формате: ФИО дд.мм.гггг" +
                    " и я покажу расписание на этот день с перерывами и обедом! ");
        } else if (txt.equals("Чекановкин Андрей")) {
            sendMsg(msg, "Приветствую, Создатель!");
        } else if (!txt.equals("")) {
            String answer = "";
            try {
                answer = readFromExcel(txt, myExcelBook);
            } catch (IOException ex){
                ex.printStackTrace();
            }
            sendMsg(msg, answer);
        }
    }

    public String getBotToken() {
        return "1144422738:AAHCxQdjQr7H0jvVOHluDHwFtYDxpaNZH1c";
    }

    private void sendMsg(Message msg, String text) {
        SendMessage message = new SendMessage();
        message.setChatId(msg.getChatId());
        message.setText(text);
        try {
            execute(message);
        } catch (TelegramApiException e) {
            e.printStackTrace();
        }
    }

    public static String readFromExcel(String message, XSSFWorkbook myExcelBook) throws IOException {
        String[] arrStr = message.split(" ");
        StringBuilder output = new StringBuilder();
        StringBuilder fio = new StringBuilder();
        ArrayList<String> table = new ArrayList<>();
        fio.append(arrStr[0]);
        fio.append(" ");
        fio.append(arrStr[1]);
        fio.append(" ");
        fio.append(arrStr[2]);

        XSSFSheet myExcelSheet = myExcelBook.getSheet("Расписание");
        Iterator<Row> rowIterator = myExcelSheet.rowIterator();
        while(rowIterator.hasNext()){
            XSSFRow row = (XSSFRow) rowIterator.next();
            if(row.getCell(3) != null){
                if(row.getCell(3).getCellType() == XSSFCell.CELL_TYPE_STRING){
                    if(row.getCell(3).getStringCellValue().equals(fio.toString())){
                        Iterator<Cell> cellIterator = myExcelSheet.getRow(2).cellIterator();
                        while(cellIterator.hasNext()){
                            XSSFCell cell = (XSSFCell) cellIterator.next();
                            if(cell != null){
                                if(cell.getCellType() == XSSFCell.CELL_TYPE_NUMERIC){
                                    Date date = cell.getDateCellValue();

                                    SimpleDateFormat format = new SimpleDateFormat("dd.MM.yyyy");
                                    String strDate = format.format(date);
                                    if(strDate.equals(arrStr[3])){
                                        for(int i = cell.getColumnIndex(); i<cell.getColumnIndex()+5; i++){
                                            String s = row.getCell(i).getRawValue();
                                            if(s.matches("(([-+])?[0-9]+(\\.[0-9]+)?)+")) {
                                                int time = (int) Math.round(row.getCell(i).getNumericCellValue() * 1440);
                                                int hours = (int) Math.floor(time / 60);
                                                int minutes = time % 60;
                                                Date data = new Date();
                                                data.setHours(hours);
                                                data.setMinutes(minutes);
                                                DateFormat workTime = new SimpleDateFormat("HH:mm");
                                                table.add(workTime.format(data));
                                            } else{
                                                table.add(s);
                                            }

                                        }
                                        myExcelBook.close();
                                    }
                                }
                            }
                        }
                    }
                }
            }
            else{
                myExcelBook.close();
            }
        }
        output.append("начало рабочего дня - ");
        output.append(table.get(0));
        output.append("\n");
        output.append("первый перерыв - ");
        output.append(table.get(1));
        output.append("\n");
        output.append("обед - ");
        output.append(table.get(2));
        output.append("\n");
        output.append("второй перерыв - ");
        output.append(table.get(3));
        output.append("\n");
        output.append("конец рабочего дня - ");
        output.append(table.get(4));
        output.append("\n");
        myExcelBook.close();
        return output.toString();
    }
}
