import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

class Main extends JFrame{
    final String DATE = "(((0[1-9]|[12]\\d|3[01])\\.(0[13578]|1[02])\\.(\\d){4})|((0[1-9]|[12]\\d|30)\\.(0[469]|11)\\.(\\d){4})|((0[1-9]|1\\d|2[0-8])\\.02\\.(\\d){4})|(29\\.02\\.(\\d){2}([02468][048]|[13579][26])))$";
    final String CELL = "[A-N]\\d+";
    final String CELL_ADD = "(\\s*\\=\\s*[A-N]\\d+\\s*([+-]\\s*\\d+)?)";
    final String DATE_ADD = "(\\s*\\=\\s*(((0[1-9]|[12]\\d|3[01])\\.(0[13578]|1[02])\\.(\\d){4})|((0[1-9]|[12]\\d|30)\\.(0[469]|11)\\.(\\d){4})|((0[1-9]|1\\d|2[0-8])\\.02\\.(\\d){4})|(29\\.02\\.(\\d){2}([02468][048]|[13579][26])))\\s*([+-]\\s*\\d+)?)";
    final String MIN = "\\s*[=]\\s*min\\s*[(]" +
            "(\\s*(((\\d\\d?\\.){2}\\d{4})|([A-Z]\\d+))\\s*[,])*\\s*(((\\d\\d?\\.){2}\\d{4})|([A-Z]\\d+))\\s*[)]";
    final String MAX = "\\s*[=]\\s*max\\s*[(]" +
            "(\\s*(((\\d\\d?\\.){2}\\d{4})|([A-Z]\\d+))\\s*[,])*\\s*(((\\d\\d?\\.){2}\\d{4})|([A-Z]\\d+))\\s*[)]";

    Main(){
        super("Excel");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setBounds(500, 100, 200, 200);
        setPreferredSize(new Dimension(700, 500));
        pack();
        JPanel panel = new JPanel(new BorderLayout());
        Object[] objects = new Object[15];
        for(int i = 0; i < 15; i++){
            objects[i] = i+1;
        }
        DefaultTableModel tableModel = new DefaultTableModel();

        tableModel.addColumn(" ",objects);
        for(int i = 0; i < 14; i++){
            tableModel.addColumn((char)('A'+i));
        }
        JTable table = new JTable(tableModel){
            @Override
            public boolean isCellEditable(int row, int column) {
                return column != 0;
            }
        };
        table.setRowHeight(48);
        table.getColumnModel().setColumnSelectionAllowed(true);
        tableModel.addTableModelListener(e->{
            int row = e.getLastRow();
            int col = e.getColumn();
            if (col > 0){
                String value = (String) tableModel.getValueAt(row, col);
                try{
                    if(Pattern.compile(CELL_ADD).matcher(value).matches()){
                    tableModel.setValueAt(cellAdd(value, tableModel), row, col);
                } else if(Pattern.compile(DATE_ADD).matcher(value).matches()){
                    tableModel.setValueAt(dateAdd(value), row, col);
                } else if(Pattern.compile(MIN).matcher(value).matches()){
                    tableModel.setValueAt(minDate(value, tableModel) ,row,col);
                } else if(Pattern.compile(MAX).matcher(value).matches()){
                    tableModel.setValueAt(maxDate(value, tableModel), row, col);
                } else throw new Exception();
                }catch(Exception exc){
                    if(!Pattern.compile(DATE).matcher(value).matches()){
                        if(!value.equals(" ")){
                            JOptionPane.showMessageDialog(null, "Wrong Input");
                            tableModel.setValueAt(" ", row, col);
                        }
                    }
                }
            }
        });
        JScrollPane scrollPane = new JScrollPane(table);
        panel.add(scrollPane, BorderLayout.CENTER);
        this.add(panel);
    }

    private Calendar toDate(String d) {
        if (d == null) return null;
        Calendar calendar;
        try {
            SimpleDateFormat simpleDateFormat = new SimpleDateFormat("dd.MM.yyyy");
            simpleDateFormat.setLenient(false);
            calendar = Calendar.getInstance();
            Date date = simpleDateFormat.parse(d);
            calendar.setTime(date);
        } catch (ParseException e) {
            return null;
        }
        return calendar;
    }

    private List<Calendar> getDateList(String d, DefaultTableModel tableModel) {
        Matcher matcher = Pattern.compile(CELL).matcher(d);
        List<Calendar> calendars = new ArrayList<>();
        while (matcher.find()) {
            String str = matcher.group();
            int col = str.charAt(0) - 'A' + 1;
            int row = Integer.parseInt(str.substring(1, str.length()))-1;
            calendars.add(toDate((String)tableModel.getValueAt(row, col)));
        }
        matcher = Pattern.compile("(\\d\\d?\\.){2}\\d{4}").matcher(d);
        while (matcher.find()) {
            calendars.add(toDate(matcher.group()));
        }
        return calendars;
    }

    private String minDate(String str, DefaultTableModel tableModel){
        Calendar calendar = Collections.min(getDateList(str, tableModel));
        return dateToString(calendar);
    }

    private String maxDate(String str, DefaultTableModel tableModel){
        Calendar calendar = Collections.max(getDateList(str, tableModel));
        return dateToString(calendar);
    }

    private String dateToString(Calendar calendar){
        int day = calendar.get(Calendar.DATE);
        int month = calendar.get(Calendar.MONTH) + 1;
        int year = calendar.get(Calendar.YEAR);
        String d = "";
        String m = "";
        if(day < 10) d += 0;
        d += day;
        if(month < 10) m += 0;
        m += month;
        return d + "." + m + "." + calendar.get(Calendar.YEAR);
    }

    private String dateAdd(String str){
        int pos = 6;
        while(Character.isDigit(str.charAt(pos)) || str.charAt(pos) == '.'){
            pos++;
        }
        String strDate = str.substring(1, pos);
        Calendar calendar = toDate(strDate);
        int add = Integer.parseInt(str.substring(pos, str.length()));
        calendar.add(GregorianCalendar.DAY_OF_MONTH, add);
        return dateToString(calendar);
    }

    private String cellAdd(String str, DefaultTableModel tableModel){
        int pos = 3;
        while(Character.isDigit(str.charAt(pos)) || Character.isLetter(str.charAt(pos))){
            pos++;
        }
        String cell = str.substring(1, pos);
        int coll = cell.charAt(0) - 'A' + 1;
        int row = Integer.parseInt(cell.substring(1, cell.length()))-1;
        SimpleDateFormat sdf = new SimpleDateFormat("dd.MM.yyyy");
        Calendar calendar = Calendar.getInstance();
        try{
            Date date = sdf.parse((String) tableModel.getValueAt(row, coll));
            calendar.setTime(date);
        } catch(ParseException e){}
        int add = Integer.parseInt(str.substring(pos, str.length()));
        calendar.add(GregorianCalendar.DAY_OF_MONTH, add);
        return dateToString(calendar);
    }

    public static void main(String[] args){
        Main main = new Main();
        main.setVisible(true);
    }
}
