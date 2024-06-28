# 关于对excel文件以及pdf文件插入页头页尾数据

## 1.在execl文件里面插入图片（文件路径/Base64字节图片）

```java
<!-- 在pom.xml文件中，引入easyExcel包 -->
<dependency>
    <groupId>com.alibaba</groupId>
    <artifactId>easyexcel</artifactId>
    <version>2.2.10</version>
</dependency>
    

<!-- 引入easyExcel包 -->    
<dependency>
    <groupId>org.apache.poi</groupId>
    <artifactId>poi</artifactId>
    <version>5.2.3</version>
</dependency>
```



```java
//	主方法调用
public static void main(String[] args) {
    // 准备数据
    List<UserData> dataList = new ArrayList<>
    // 写入数据到 Excel
    String filePath = "example.xlsx";
    EasyExcel.write(filePath, null).sheet("Sheet1").doWrite(dataLis
    // 插入图片
    insertImage(filePath, "Snipaste_2024-06-13_17-06-06.png", 1, 4);  // 替换为你的图片路径
	// 插入图片（使用 Base64 字符串）
     String base64String = "";  // 替换为你的 Base64 字符串
    insertImageBase64(filePath, base64String, 1, 4);
}

/**
 * @Description: 插入图片方法
 * @param excelFilePath 文件的路径
 * @param base64String  图片的base64字节数
 * @param row excel文件的位置 行位置
 * @param col excel文件的位置 列位置
 * @return: void
 */
public static void insertImageBase64(String excelFilePath, String base64String, int row, int col) {
        try (FileInputStream excelFile = new FileInputStream(excelFilePath);
             Workbook workbook = new XSSFWorkbook(excelFile)) {

            // 解码 Base64 字符串为字节数组
            byte[] imageBytes = Base64.getDecoder().decode(base64String);

            int pictureIdx = workbook.addPicture(imageBytes, Workbook.PICTURE_TYPE_PNG);

            // 获取第一个工作表
            Sheet sheet = workbook.getSheetAt(0);

            // 创建绘图对象
            Drawing<?> drawing = sheet.createDrawingPatriarch();

            // 设置图片插入位置和大小
            ClientAnchor anchor = workbook.getCreationHelper().createClientAnchor();
            anchor.setCol1(col);  // 图片起始列
            anchor.setRow1(row);  // 图片起始行
            anchor.setCol2(col + 2);  // 图片结束列
            anchor.setRow2(row + 2);  // 图片结束行

            // 插入图片
            drawing.createPicture(anchor, pictureIdx);

            // 保存修改后的文件
            try (FileOutputStream outFile = new FileOutputStream(excelFilePath)) {
                workbook.write(outFile);
            }

            System.out.println("图片插入成功！");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

                                                            
    /**
     * @Description: 插入图片方法
     * @param excelFilePath 文件的路径
     * @param imagePath 图片的路径
     * @param row excel文件的位置 行位置
     * @param col excel文件的位置 列位置
     * @return: void
     */
public static void insertImage(String excelFilePath, String imagePath, int row, int col) {
        try (FileInputStream excelFile = new FileInputStream(excelFilePath);
             Workbook workbook = new XSSFWorkbook(excelFile);
             FileInputStream imageFile = new FileInputStream(imagePath)) {

            byte[] bytes = IOUtils.toByteArray(imageFile);
            int pictureIdx = workbook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);

            Sheet sheet = workbook.getSheetAt(0);
            Drawing<?> drawing = sheet.createDrawingPatriarch();
            ClientAnchor anchor = workbook.getCreationHelper().createClientAnchor();
            anchor.setCol1(col);  // 图片插入位置的列索引
            anchor.setRow1(row);  // 图片插入位置的行索引
            anchor.setCol2(col + 2);  // 图片结束列
            anchor.setRow2(row + 2);  // 图片结束行
            drawing.createPicture(anchor, pictureIdx);

            // 重新保存文件
            try (FileOutputStream outFile = new FileOutputStream(excelFilePath)) {
                workbook.write(outFile);
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
}
```

### 2.获取excel里面的单个图片（目前是只存了一个图片）

```java
    public static void main(String[] args) throws Exception {

        File file = new File("文件地址.xlsx");

        if (!file.exists() || !file.isFile()) {
            System.out.println("Invalid file path provided.");
            return;
        }

        try (InputStream inputStream = new FileInputStream(file)) {
            boolean hasImage = containsImage(inputStream,"图片数据的文件路径");
            System.out.println("Excel文件里面包含了图片: " + hasImage);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

/**
 * @Description: 获取excel文件里面的图片
 * @param inputStream excel文件流
 * @param outPutFilePath 图片输出的路径
 * @return: boolean 是否获取到了文件
 */
private static boolean containsImage(InputStream inputStream, String outPutFilePath) throws Exception{
    Workbook workbook = WorkbookFactory.create(inputStream);

    for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
        Sheet sheet = workbook.getSheetAt(i);
        Drawing<?> drawing = sheet.getDrawingPatriarch();

        if (drawing != null) {
            List<? extends PictureData> pictures = workbook.getAllPictures();
            PictureData pictureData = pictures.get(0);
            //  后缀
            String ext = pictureData.suggestFileExtension();
            //  图片数据
            byte[] data = pictureData.getData();
            File dir = new File(outPutFilePath);
            if (!dir.exists()) {
                dir.mkdirs();
            }
            File imgFile = new File(dir, "image_" + System.currentTimeMillis() + "." + ext);
            try (FileOutputStream out = new FileOutputStream(imgFile)) {
                out.write(data);
            }
            if (!pictures.isEmpty()) {
                return true;
            }
        }
    }
    return false;
}
```

### 3.在excel里面的固定文字的后面插入图片

```java
public static void main(String[] args) {

        String filePath = "清单文件.xlsx";
        String searchString = "张三";

        File file = new File(filePath);

        if (!file.exists() || !file.isFile()) {
            System.out.println("Invalid file path provided.");
            return;
        }

        try (InputStream inputStream = new FileInputStream(file)) {
            findString(inputStream, searchString);
        } catch (IOException | InvalidFormatException e) {
            e.printStackTrace();
        }
    }

/**
 * @Description: 查找这个文件的位置，在图片后面的一个单元格插入图片
 * @param inputStream excel文件流
 * @param searchString 需要搜索的文字
 * @return: void
 */    
public static void findString(InputStream inputStream, String searchString) throws IOException, InvalidFormatException {
        Workbook workbook = WorkbookFactory.create(inputStream);
        String filePath1 = "清单文件.xlsx";
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            Sheet sheet = workbook.getSheetAt(i);
            for (Row row : sheet) {
                for (Cell cell : row) {
                    if (cell.getCellType() == CellType.STRING) {
                        String cellValue = cell.getStringCellValue();
                        if (cellValue != null && cellValue.contains(searchString)) {
                            //  插入图片
                             // 替换为你的图片路径
                            //	调用插入图片的方法
                            insertImage(filePath1, "Snipaste_2024-06-13_17-06-06.png", row.getRowNum(),cell.getColumnIndex()+1); 
                            System.out.println("Found '" + searchString + "' at sheet " + sheet.getSheetName() + ", row " + row.getRowNum() + ", column " + cell.getColumnIndex());
                        }
                    }
                }
            }
        }
    }
```

