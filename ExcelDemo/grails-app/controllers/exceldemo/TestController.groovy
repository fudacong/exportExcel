package exceldemo

import org.apache.poi.hssf.usermodel.HSSFCell
import org.apache.poi.hssf.usermodel.HSSFRow
import org.apache.poi.hssf.usermodel.HSSFSheet
import org.apache.poi.hssf.usermodel.HSSFWorkbook
import java.lang.reflect.Field

class TestController {

    /*调远程数据的接口*/
    RemoteUserService remoteUserService

    def index() {

    }

    /*excel导出*/
    def excelPort(){

        /*创建一个工作薄,在此工作薄下创建一个sheet1页*/
        HSSFWorkbook workbook = new HSSFWorkbook();                 //声明一个工作薄,就是创建一个excel文件
        HSSFSheet sheet = workbook.createSheet("sheet1");           //声明一个sheet页,并命名标题
        sheet.setDefaultColumnWidth(20);                             //设置单元格默认宽度

        /*标题行和属性名的对应*/
        Map<String,String> titleKeyMap = [
                "用户名": "userName",
                "密码"  : "password",
                "性别"  : "sex",
                "年龄"  : "age",
                "爱好"  : "hobby",
                "公司"  : "company"
        ]

        /*取出标题*/
        List<String> titles = new ArrayList<String>(titleKeyMap.keySet())

        /*行号索引*/
        int rowIndex = 0;

        /*创建标题行*/
        HSSFRow row1 = sheet.createRow(rowIndex ++);
        row1.setHeightInPoints((short)30);                              //设置标题行的高度为30
        for(int i = 0 ; i < titles.size(); i++){
            HSSFCell cell = row1.createCell(i);                        //给该行创建单元格
            cell.setCellValue(titles.get(i))                           //给单元格放入标题
        }

        /*从数据库中查出列表数据(造假数据)*/
        List<UserBean> userBeanList = remoteUserService.findUserList()

        /*根据数据的条数来创建表格行*/
        for(UserBean userBean : userBeanList){
            HSSFRow row = sheet.createRow(rowIndex ++ )
            for( int k = 0 ; k < titles.size() ; k ++){
                HSSFCell cell = row.createCell(k);
                String propertyName = titleKeyMap.get(titles[k]);
                Field field = userBean.getClass().getDeclaredField(propertyName)         //利用反射根据属性名获取属性值
                field.setAccessible(true);                                               //强制访问私有属性
                String value = String.valueOf( field.get(userBean))                      //从userBean中取出该属性的值
                cell.setCellValue(value)
            }
        }

        //提示用户下载该excel文档
        try
        {
            response.setContentType("text/html;charset=utf-8");                                 //设置响应编码
            response.setContentType("application/x-msdownload");                                //设置为文件下载
            response.addHeader("Content-Disposition", "attachment;filename=myExcel.xls");      //设置相应头信息和默认文件名
            OutputStream outputStream = response.getOutputStream();                             //创建输出流
            workbook.write(outputStream);                                                       //把工作薄写进流中
            outputStream.close();
        }
        catch (IOException e)
        {
            e.printStackTrace();
        }

    }

}
