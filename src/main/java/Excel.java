import java.io.*;
import java.sql.*;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import com.google.gson.*;


public class Excel {
//    public static final String xlsxPath = "./xiaoqurelations(导入)-20190516(给赵旺).xlsx";
//    public static final String xlsxPath = "./xiaoqurelations_update(20190606).xlsx";
//    public static final String xlsxPath = "./xiaoqurelations(导入)-20190610.xlsx";
    public static String xlsxPath = "";

    public static String MOVE = "";
    //    public static final String USERNAME = "zhuchenxiao";
//    public static final String USERNAME = "root";
    public static String USERNAME = "";

//    public static final String PASSWD = "HPFbLZsQZ9MRr1dT92NH";
//    public static final String PASSWD = "werwer";
    public static String PASSWD = "";

//        public static final String DBURL = "jdbc:mysql://rr-2zetk6oj0m9p6g95j.mysql.rds.aliyuncs.com/xiaoqurelations?useUnicode=true&characterEncoding=utf-8";
//    public static final String DBURL = "jdbc:mysql://localhost/xiaoqu?useUnicode=true&characterEncoding=utf-8";
    public static String DBURL = "";

//    public static final String TABLE_NAME = "xiaoqurelations_0610";
    public static String TABLE_NAME = "";

    public static final int COLUMNS = 10;
    public static int DELETE_SHEET_INDEX = 0;

//    public static final int INPUT_SHEET_INDEX = 1;
    public static int INPUT_SHEET_INDEX = 0;

    // 构造方法
    public Excel(){
        readJson();
    }

    //*******************************读取Json中的配置******************************************
    public void readJson(){
        try{
            String path = "./excelinfo.json";
            BufferedReader bufferedReader = new BufferedReader(new FileReader(path));
            Gson gson = new Gson();
            JsonObject js = gson.fromJson(bufferedReader, JsonObject.class);
            MOVE = js.get("move").getAsString();
            xlsxPath = js.get("xlsxPath").getAsString();
            USERNAME = js.get("db_username").getAsString();
            PASSWD = js.get("db_password").getAsString();
            DBURL = js.get("db_url").getAsString();
            TABLE_NAME = js.get("db_tablename").getAsString();
            DELETE_SHEET_INDEX = js.get("delete_sheet_index").getAsInt();
            INPUT_SHEET_INDEX = js.get("input_sheet_index").getAsInt();
            System.out.println(MOVE+"\n"+xlsxPath+"\n"+USERNAME+"\n"+PASSWD+"\n"+DBURL+"\n"+TABLE_NAME+"\n"+DELETE_SHEET_INDEX+"\n"+INPUT_SHEET_INDEX);
        }catch (Exception e){
            e.printStackTrace();
        }
    }

    //***************************************************************************************

    //*******************************Excel提取两个Sheet******************************************
    public List<XiaoQuRelations> delete_ExcelInfo() throws IOException {
        // 将excel中的要删除数据添加到list中返回
        List temp = new ArrayList();
        FileInputStream fileIn = new FileInputStream(xlsxPath);
        //根据指定的文件输入流导入Excel从而产生XSSFWorkbook对象
        XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(xlsxPath));
        //获取Excel文档中的第一个表单：删除
        Sheet sht_delete = xwb.getSheetAt(DELETE_SHEET_INDEX);

        //对Sheet中的每一行进行迭代
        for (Row r : sht_delete) {
            //如果当前行的行号（从0开始）未达到2（第三行）则从新循环
            if(r.getRowNum()<1){
                continue;
            }
            //创建实体类
            XiaoQuRelations info=new XiaoQuRelations();
            //取出单元格数据，并封装在info实体属性
            for (int i=0;i<COLUMNS;i++){
                if (r.getCell(i) != null){
                    switch (i){
                        case 0:
                            info.setId((int)r.getCell(0).getNumericCellValue());
                            break;
                        case 1:
                            info.setCity(r.getCell(1).getStringCellValue());
                            break;
                        case 2:
                            info.setDistrict(r.getCell(2).getStringCellValue());
                            break;
                        case 3:
                            info.setBlock(r.getCell(3).getStringCellValue());
                            break;
                        case 4:
                            info.setType(r.getCell(4).getStringCellValue());
                            break;
                        case 5:
                            info.setCompany(r.getCell(5).getStringCellValue());
                            break;
                        case 6:
                            info.setJingdui(r.getCell(6).getStringCellValue());
                            break;
                        case 7:
                            info.setNote(r.getCell(7).getStringCellValue());
                            break;
                        case 8:
                            info.setJingdui_type(r.getCell(8).getStringCellValue());
                            break;
                        case 9:
                            info.setJingdui_district(r.getCell(9).getStringCellValue());
                            break;
                    }
                }
            }
            temp.add(info);
//            info.show();
//            break;
        }
        fileIn.close();
        return temp;
    }

    public  List<XiaoQuRelations> input_ExcelInfo() throws IOException{
        // 将excel中的要删除数据添加到list中返回
        List temp = new ArrayList();
        FileInputStream fileIn = new FileInputStream(xlsxPath);
        //根据指定的文件输入流导入Excel从而产生XSSFWorkbook对象
        XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(xlsxPath));
        //获取第二个表单：增加
        Sheet sht_input = xwb.getSheetAt(INPUT_SHEET_INDEX);
        //对Sheet中的每一行进行迭代
        for (Row r : sht_input) {
            //如果当前行的行号（从0开始）未达到2（第三行）则从新循环
            if(r.getRowNum()<1){
                continue;
            }
            //创建实体类
            XiaoQuRelations info=new XiaoQuRelations();
            //取出单元格数据，并封装在info实体stuName属性
            for (int i=0;i<COLUMNS;i++){
                if (r.getCell(i) != null){
                    switch (i){
                        case 0:
                            info.setId((int)r.getCell(0).getNumericCellValue());
                            break;
                        case 1:
                            info.setCity(r.getCell(1).getStringCellValue());
                            break;
                        case 2:
                            info.setDistrict(r.getCell(2).getStringCellValue());
                            break;
                        case 3:
                            info.setBlock(r.getCell(3).getStringCellValue());
                            break;
                        case 4:
                            info.setType(r.getCell(4).getStringCellValue());
                            break;
                        case 5:
                            info.setCompany(r.getCell(5).getStringCellValue());
                            break;
                        case 6:
                            info.setJingdui(r.getCell(6).getStringCellValue());
                            break;
                        case 7:
                            info.setNote(r.getCell(7).getStringCellValue());
                            break;
                        case 8:
                            info.setJingdui_type(r.getCell(8).getStringCellValue());
                            break;
                        case 9:
                            info.setJingdui_district(r.getCell(9).getStringCellValue());
                            break;
                    }
                }
            }
            temp.add(info);
//            info.show();
//            break;
        }
        fileIn.close();
        return temp;
    }

    //***************************************************************************************



    //*******************************数据库连接部分******************************************
    public static Connection connectDB() throws ClassNotFoundException, SQLException{
        Class.forName("com.mysql.jdbc.Driver");
        Connection conn=DriverManager.getConnection(DBURL,USERNAME,PASSWD);
        System.out.println("[数据库连接成功...]");
        return conn;
    }
    // 关闭数据库连接
    public static void closeAll(Connection conn, PreparedStatement pst, ResultSet rs){
        try {
            if(rs!=null){
                rs.close();
            }
            if(pst!=null){
                pst.close();
            }
            if(conn!=null){
                conn.close();
            }
            System.out.println("[数据库关闭...]");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    //***************************************************************************************


    //********************************删除与增加部分***********************************************
    public void deleteFromDB(List<XiaoQuRelations> list) throws ClassNotFoundException,SQLException{
        // 查询和删除代码分开
//        String sql_delete = "DELETE FROM "+TABLE_NAME+" WHERE id = ? AND city = ? AND district = ? AND block = ? "
//              + "AND type = ? AND company = ? AND jingdui = ? AND jingdui_type = ? AND jingdui_district = ?;";
        String sql_delete_ijj = String.format("DELETE FROM %s WHERE id = ? AND jingdui = ? AND jingdui_type = ? ;", TABLE_NAME);

//        String sql_query = "SELECT * FROM "+TABLE_NAME+" WHERE id = ? AND city = ? AND district = ? AND block = ? "
//                + "AND type = ? AND company = ? AND jingdui = ? AND  jingdui_type = ? AND jingdui_district = ?;";
//        String sql_query_ijj = String.format("SELECT * FROM %s WHERE id = ? AND jingdui = ? AND jingdui_type = ? ;", TABLE_NAME);
        String sql_query_ijj = String.format("SELECT id,jingdui,jingdui_type FROM %s WHERE id = ? AND jingdui = ? AND jingdui_type = ? ;", TABLE_NAME);

        PreparedStatement pst_delete = connectDB().prepareStatement(sql_delete_ijj);
        PreparedStatement pst_query = connectDB().prepareStatement(sql_query_ijj);
        int affected_rows = 0;
//        int i = 0;
        // 遍历表格中的删除数据
        for (XiaoQuRelations xiaoqu : list) {
//            System.out.println(pst_query.toString());
//            ResultSet rs = pst_query.executeQuery();
            pst_query.setString(1,xiaoqu.getId()+"");
            pst_query.setString(2,xiaoqu.getJingdui());
            pst_query.setString(3,xiaoqu.getJingdui_type());
            ResultSet rs = pst_query.executeQuery();
            boolean result = rs.next();
            System.out.println("查询结果："+xiaoqu.getId()+" "+result);
//             如果查得到就删除,查不到则不处理
//            if (false) {
            if (result == true) {
                System.out.println("找到id为"+xiaoqu.getId()+"的数据，正在删除...");
                pst_delete.setString(1, xiaoqu.getId() + "");
                pst_delete.setString(2, xiaoqu.getJingdui());
                pst_delete.setString(3, xiaoqu.getJingdui_type());
//                System.out.println(pst_delete.toString());
                int row = pst_delete.executeUpdate();
                if(row!=0){
                    affected_rows += row;
                }
                // rows affected
                System.out.println("-----成功删除id为"+xiaoqu.getId()+"的数据！");
                System.out.println("Rows Affected:"+row);
            }
        }
        System.out.println("数据库删除完成！");
        System.out.println("Rows Totally Affected:"+affected_rows);
    }


    public void inputFromDB(List<XiaoQuRelations> list) throws ClassNotFoundException,SQLException{
    // 查询和添加代码分开
        String sql_query = String.format("SELECT id,jingdui,jingdui_type FROM %s WHERE id = ? AND jingdui = ? AND jingdui_type = ? ;", TABLE_NAME);
        String sql_insert = String.format("INSERT INTO %s VALUES(?,?,?,?,?,?,?,?,?,?)", TABLE_NAME);
        PreparedStatement pst_query = connectDB().prepareStatement(sql_query);
        PreparedStatement pst_insert = connectDB().prepareStatement(sql_insert);
        int affected_rows = 0;
        // 遍历表格中的添加数据
        for (XiaoQuRelations xiaoqu : list) {
            pst_query.setString(1,xiaoqu.getId()+"");
            pst_query.setString(2,xiaoqu.getJingdui());
            pst_query.setString(3,xiaoqu.getJingdui_type());
            ResultSet rs = pst_query.executeQuery();
            Boolean result = rs.next();
            System.out.println("查询结果："+result);
            // 如果查得到就不处理,查不到则添加
            if (result==false) {
//            if (false) {
                pst_insert.setString(1, xiaoqu.getId() + "");
                System.out.println(xiaoqu.getCity()+" "+xiaoqu.getDistrict());
                pst_insert.setString(2, xiaoqu.getCity());
                pst_insert.setString(3, xiaoqu.getDistrict());
                pst_insert.setString(4, xiaoqu.getBlock());
                pst_insert.setString(5, xiaoqu.getType());
                pst_insert.setString(6, xiaoqu.getCompany());
                pst_insert.setString(7, xiaoqu.getJingdui());
                pst_insert.setString(8, xiaoqu.getNote());
                pst_insert.setString(9, xiaoqu.getJingdui_type());
                pst_insert.setString(10, xiaoqu.getJingdui_district());
                int row = pst_insert.executeUpdate();
                if(row==1){
                    affected_rows += row;
                }
                // rows affected
                System.out.println("-----成功添加id为"+xiaoqu.getId()+"的数据！");
                System.out.println("Rows Affected:"+row);
            }
        }
        System.out.println("数据库添加完成！");
        System.out.println("Rows Totally Affected:"+affected_rows);
    }

    public static void test(String id)throws ClassNotFoundException,SQLException{

        String sql_query = String.format("SELECT * FROM %s WHERE id = ?;", TABLE_NAME);
        PreparedStatement pst_query = connectDB().prepareStatement(sql_query);
        pst_query.setString(1,id);
//        pst_query.setString(2,id);
        System.out.println(pst_query.toString());
        System.out.println(id);
        ResultSet rs = pst_query.executeQuery();
        System.out.println("找的数据信息如下:");
        while(rs.next()){
            System.out.println(rs.getString("id")+" "+rs.getString("city")+" "+rs.getString("district")+" "+rs.getString("block")
                    +" "+rs.getString("type")+" "+rs.getString("company")+" "+rs.getString("jingdui")+" "+rs.getString("note")
                    +" "+rs.getString("jingdui_type")+" "+rs.getString("jingdui_district"));
        }
    }
    //***************************************************************************************
    public static void main(String[] args){
        long start=System.currentTimeMillis();
        Excel ex = new Excel();
        // -d 代表删除 -i代表添加 -t表示测试
//        if (args[0].equals("-d")) {
        if (ex.MOVE.equals("-d")) {
            // 执行删除操作
            try {
                List<XiaoQuRelations>  delete_lst = ex.delete_ExcelInfo();
                System.out.println("要删除的数据总量："+delete_lst.size());
                ex.deleteFromDB(delete_lst);
            } catch (Exception e) {
            e.printStackTrace();
            }
        }
        else if (ex.MOVE.equals("-i")){
            // 执行添加操作
            try {
                List<XiaoQuRelations>  input_lst = ex.input_ExcelInfo();
                System.out.println("要添加的数据总量："+input_lst.size());
                ex.inputFromDB(input_lst);

            } catch (Exception e) {
                e.printStackTrace();
            }
        }
        else if (ex.MOVE.equals("-di") || ex.MOVE.equals("-id")){
            // 执行添加+删除
            try {
                List<XiaoQuRelations>  delete_lst = ex.delete_ExcelInfo();
                System.out.println("要删除的数据总量："+delete_lst.size());
                ex.deleteFromDB(delete_lst);
                System.out.println("删除数据结束，即将开始添加数据...");
                List<XiaoQuRelations>  input_lst = ex.input_ExcelInfo();
                System.out.println("要添加的数据总量："+input_lst.size());
                ex.inputFromDB(input_lst);

            } catch (Exception e) {
                e.printStackTrace();
            }
        }
        else if(ex.MOVE.equals("-t")){
            // 执行测试，删除或者添加成功与否，使用select查找id
            try{
                //输入测试id
                Excel.test("41231");
            }catch(Exception e){e.printStackTrace();}
        }else{
            System.out.println("请正确输入参数！-d表示删除，-i表示添加");
        }
        Long end=System.currentTimeMillis();
        System.out.println("用时："+(end-start));
    }
}
