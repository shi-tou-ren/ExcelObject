# ExcelObject
简化Excel文件创建，并可持续化对Excel文件进行更新的项目

---
### 文档
[点击查看文档](https://shi-tou-ren.github.io/ExcelObject/index.html)

---
### Apache Maven


```xml
<dependency>
  <groupId>com.github.shi-tou-ren</groupId>
  <artifactId>eo</artifactId>
  <version>1.0.t</version>
</dependency>
```

---
### 简单使用


```
public class Test {

    private Integer id;

    private String name;

    private Date time;

    //简单使用
    public static void main(String[] args) throws FileNotFoundException {
        //创建空白excel文件
        Excel excel = new Excel();
        //通过配置创建空白表格
        Excel.Sheet sheet = excel.createSheet("id:id,name:名称,time:时间");
        //生成表格数据1
        List<Test> testList1 = new ArrayList<Test>();
        Test test1 = new Test();
        test1.setId(1);
        test1.setName("Tony");
        test1.setTime(new Date());
        Test test2 = new Test();
        test2.setId(2);
        test2.setName("Mary");
        test2.setTime(new Date());
        testList1.add(test1);
        testList1.add(test2);
        //生成表格数据2
        List<Test> testList2 = new ArrayList<Test>();
        Test test3 = new Test();
        test3.setId(3);
        test3.setName("Allen");
        test3.setTime(new Date());
        Test test4 = new Test();
        test4.setId(4);
        test4.setName("Olivia");
        test4.setTime(new Date());
        testList2.add(test3);
        testList2.add(test4);
        //更改表格名称
        sheet.rename("测试表格1");
        //持续多次添加表格数据
        sheet.addData(testList1);
        sheet.addData(testList2);
        //通过IO流导出表格 仅支持xls格式的文件
        OutputStream outputStream = new FileOutputStream(new File("D:\\test.xls"));
        excel.export(outputStream);
    }
    
    public Integer getId() {
        return id;
    }

    public void setId(Integer id) {
        this.id = id;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Date getTime() {
        return time;
    }

    public void setTime(Date time) {
        this.time = time;
    }

}

```

---
### 配置文件
配置文件格式为:  
- {成员属性名称:对应的表格列名称,成员属性名称:对应的表格列名称}
- 成员属性名称:对应的表格列名称,成员属性名称:对应的表格列名称
 


例：
- {id:id,name:名称,time:时间}
- id:id,name:名称,time:时间

---
### 注意事项
1: 导出的文件格式请设置为.xls格式；