# 使用指南
 只需要修改excelinfo.json中的配置即可，其中字段的含义如下：
 
 
 - move 
 
 代表要执行的操作，-d代表删除，-i代表添加，-di或者-id代表删除+添加
 
 - xlsxPath
 
 代表要读取的excel文件的位置，需要放在Excel目录下才能读取到，不知道为什么别的位置读取不到。
 
 - db_username
 
 代表数据库的用户名
 
 - db_password
 
 代表数据库用户的密码
 
 - db_url
 
 代表数据库的地址，注意要加上库名称，实例如下：
 
 "jdbc:mysql://localhost/xiaoqu?useUnicode=true&characterEncoding=utf-8" 
 
 ？后面的部分代表编码方式，使用utf8可以避免中文乱码
 
 - db_tablename
 
 代表要修改的表的名字
 
 - delete_sheet_index
 
 代表在excel中要删除部分的表格的索引，默认为0，可以指定sheet
 
 - input_sheet_index
 
 代表excel中要添加部分的表格的索引，默认为0，可以指定sheet
 
 
 
 # 注意测试
 
 由于删除操作风险比较大，在deleteFromDB中将删除部分做了一些可测试，if(false)时删除部分不会执行，只会打印出查询导的要删除的数据信息，if(result == true)时执行删除
 
 同理处理了inputFromDB函数

 2019-06-20 优化了查询语句，数据库添加多列索引，提升查找速度。