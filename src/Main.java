import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

class Main extends JFrame{

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
        final String DATE = "(((0[1-9]|[12]\\d|3[01])\\.(0[13578]|1[02])\\.(\\d){4})|((0[1-9]|[12]\\d|30)\\.(0[469]|11)\\.(\\d){4})|((0[1-9]|1\\d|2[0-8])\\.02\\.(\\d){4})|(29\\.02\\.(\\d){2}([02468][048]|[13579][26])))[ \n]?";
        final String CELL = "(\\s*\\=\\s*[A-Z]\\d+\\s*([+-]\\s*\\d+)?)";
        Pattern pattern = Pattern.compile(DATE);
        tableModel.addTableModelListener(e->{
            int row = e.getLastRow();
            int col = e.getColumn();
            if (col > 0){
                try{
                String value = (String) tableModel.getValueAt(row, col);
                    if(value.charAt(0) == '='){
                        if(value.contains("+") && value.indexOf("+") > 1){
                            int i = value.indexOf("+");
                            Matcher matcher = pattern.matcher(value.substring(1, i));
                            if(Pattern.compile(DATE).matcher(value.substring(1,i)).matches() && i < value.length()-1){
                                Matcher m = Pattern.compile("[+]?[1-9]\\d*").matcher(value);
                                ArrayList<Integer> arr = new ArrayList<>(4);
                                while(m.find()){
                                    arr.add(Integer.parseInt(m.group()));
                                }
                                Calendar calendar;
                                SimpleDateFormat simpleDateFormat = new SimpleDateFormat("dd.MM.yyyy");
                                simpleDateFormat.setLenient(false);
                                calendar = Calendar.getInstance();
                                Date date = simpleDateFormat.parse(value.substring(1, i));
                                calendar.setTime(date);
                                calendar.add(GregorianCalendar.DAY_OF_MONTH, arr.get(3));
                                int day = calendar.get(Calendar.DATE);
                                int month = calendar.get(Calendar.MONTH)+1;
                                int year = calendar.get(Calendar.YEAR);
                                String d, mh, y;
                                if(day < 10) d = 0+""+day;
                                else d = ""+day;
                                if(month < 10) mh = 0+""+month;
                                else mh = ""+month;
                                y = ""+year;
                                tableModel.setValueAt(d+"."+mh+"."+y, row, col);
                            } else if(Pattern.compile(CELL).matcher(value).matches()){
                                int a = value.charAt(1)-'A'+1;
                                int b = Integer.parseInt(value.substring(2, value.indexOf("+")))-1;
                                Calendar calendar;
                                SimpleDateFormat simpleDateFormat = new SimpleDateFormat("dd.MM.yyyy");
                                simpleDateFormat.setLenient(false);
                                calendar = Calendar.getInstance();
                                String s = (String)tableModel.getValueAt(b, a);
                                Date date = simpleDateFormat.parse(s);
                                calendar.setTime(date);
                                calendar.add(GregorianCalendar.DAY_OF_MONTH, Integer.parseInt(value.substring(value.indexOf("+")+1, value.length())));
                                int day = calendar.get(Calendar.DATE);
                                int month = calendar.get(Calendar.MONTH)+1;
                                int year = calendar.get(Calendar.YEAR);
                                String d, mh, y;
                                if(day < 10) d = 0+""+day;
                                else d = ""+day;
                                if(month < 10) mh = 0+""+month;
                                else mh = ""+month;
                                y = ""+year;
                                tableModel.setValueAt(d+"."+mh+"."+y, row, col);
                            }
                        } else if(value.contains("-") && value.indexOf("-") > 1){

                        } else if(!Pattern.compile(DATE).matcher(value).matches()){
                            throw new Exception();
                        }
                    } else throw new Exception();
                }catch(Exception exc){
                    //JOptionPane.showMessageDialog(null, "Error");
                }
            }
        });
        JScrollPane scrollPane = new JScrollPane(table);
        panel.add(scrollPane, BorderLayout.CENTER);
        this.add(panel);
    }

    public static void main(String[] args){
        Main main = new Main();
        main.setVisible(true);
    }
}