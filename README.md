# jexcel
基于 Apache POI 开发的 Excel 文件读写工具包

# 样例

定义一个简单 POJO，并用 **@Cell** 注解标注数据字段：

```java
public class User {
 
     @Cell(index = 0, name = "编号", align = Align.CENTER)
     private int id;
 
     @Cell(index = 1, name = "姓名", align = Align.CENTER)
     private String name;
 
     @Cell(index = 2, name = "日期", align = Align.CENTER)
     private Date date;
 
     public int getId() {
         return id;
     }
 
     public void setId(int id) {
         this.id = id;
     }
 
     public String getName() {
         return name;
     }
 
     public void setName(String name) {
         this.name = name;
     }
 
     public Date getDate() {
         return date;
     }
 
     public void setDate(Date date) {
         this.date = date;
     }
 
 }
 ````
 
 ## 生成 Excel 文件
 
 ------------------
 
 ```java
public static void main(String[] args) {
    ExcelExecutor.write(getUsers()).toFile("D:\\test.xlsx");
}

// 模拟数据
private static List<User> getUsers() {
    List<User> users = new ArrayList<>();
    for (int i = 1; i <= 20; i++) {
        User user = new User();
        user.setId(i);
        user.setDate(new Date());
        user.setName("U" + ThreadLocalRandom.current().nextInt(100));
        users.add(user);
    }
    return users;
}
```
 
 ## 生成 Excel 文件
 
 ------------------