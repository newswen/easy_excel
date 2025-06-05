package com.yw.easy_excel_test.utils;
 
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
 
import javax.imageio.ImageIO;
import javax.swing.*;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.*;
 
public class InsertTxtFileAsOleObject {
 
 
    public static void main(String[] args) throws IOException {
        //只有HSSFWorkbook才能使用OLE对象，并且poi需要在4.0之上
		// 目标Excel文件路径：若为 .xlsx，将在同目录生成/更新同名 .xls（CS.xls），并在其中嵌入 OLE
		String excelPath = "D:\\xfg\\easy_excel_test\\src\\main\\java\\com\\yw\\easy_excel_test\\utils\\CS.xls";
		String lower = excelPath.toLowerCase();
		String outputPath = lower.endsWith(".xlsx") ? excelPath.substring(0, excelPath.length() - 5) + ".xls" : excelPath;

		HSSFWorkbook workbook;
		File outFile = new File(outputPath);
		if (outFile.exists()) {
			try (FileInputStream excelFis = new FileInputStream(outFile)) {
				workbook = new HSSFWorkbook(excelFis);
			}
		} else if (!lower.endsWith(".xlsx") && new File(excelPath).exists()) {
			try (FileInputStream excelFis = new FileInputStream(excelPath)) {
				workbook = new HSSFWorkbook(excelFis);
			}
		} else {
			workbook = new HSSFWorkbook();
		}
		Sheet sheet = workbook.getSheet("Sheet1");
		if (sheet == null) {
			sheet = workbook.createSheet("Sheet1");
		}
 
 
        /** 固定展示区域：1 行 3 列，固定像素大小 */
        int startCol = 2; // 从第3列开始（C列）
        int startRow = 0; // 第1行
        int colsSpan = 3; // 横跨3列
        int rowsSpan = 1; // 占1行
        int totalWidthPx = 180; // 固定总宽（像素）——可按需调整
        int heightPx = 60; // 固定高度（像素）——可按需调整

        // 设置这3列的固定宽度
        int perColPx = totalWidthPx / colsSpan;
        for (int i = 0; i < colsSpan; i++) {
            sheet.setColumnWidth(startCol + i, pixelsToColumnWidthUnits(perColPx));
        }
        // 设置这一行的固定高度
        Row row = sheet.getRow(startRow);
        if (row == null) {
            row = sheet.createRow(startRow);
        }
        row.setHeightInPoints(pixelsToPoints(heightPx));
 
 
        // 读取需要添加的文件，如果有需要添加多个文件的需求，可以循环表格来添加
        File oleFile = new File("D:\\Dsesktop\\test.docx");
        FileInputStream fis = new FileInputStream(oleFile);
        byte[] fileBytes = new byte[(int) oleFile.length()];
        fis.read(fileBytes);
        fis.close();
 
//        /** 这一步是获取文件图标 */
//        // 获取文件系统视图
//        FileSystemView view = FileSystemView.getFileSystemView();
//        Icon systemIcon = view.getSystemIcon(oleFile);
//        // 将 Icon 对象转换为 Image 对象
//        Image image = iconToImage(systemIcon);
//        // 将 Image 对象转换为字节数组（缩放到固定尺寸，避免初始显示被裁剪或双击后突变）
//        byte[] imageBytes = imageToByteArray(image, 20, 20);
 
 
        /**  自定义文件的展示图标 */
         String imagePath = "D:\\xfg\\easy_excel_test\\src\\main\\java\\com\\yw\\easy_excel_test\\utils\\img.png"; // 图片文件路径 这个是excel中单元格中pdf文件对应展示出来的图片，可以自定义
         BufferedImage image = ImageIO.read(new File(imagePath));
         ByteArrayOutputStream baos = new ByteArrayOutputStream();
         ImageIO.write(image, "png", baos);
         byte[] imageBytes = baos.toByteArray();
         baos.close();
 
 
        //将文件的图标添加进入到Excel文件内
        int iconid = workbook.addPicture(imageBytes, Workbook.PICTURE_TYPE_PNG);
        // 在工作表中创建OLE对象，就是将文件插入到Excel文件中
        int oleIdx = workbook.addOlePackage(fileBytes, "test.docx", "test.docx", "test.docx");
 
         /**
          pdfBytes: 表示 PDF 文件的字节数组。您需要将 PDF 文件的内容以字节数组的形式传递给此参数。
          "222.zip": 表示 OLE 对象的类型。如果导出之后，在excel中操作另存文件，这个名称就是新保存文件的默认名称【文件标签】
          "333.zip": 这指定了要在OLE Package中使用的类名。在此示例中，你可以将其设置为任何你想要的类名，因为它不会对后续的操作产生影响。它也只是用于标识特定的OLE Package【文件名】。
          "111.zip":这是在Excel工作簿中显示的OLE Package的名称,如果需要在excel中打开pdf文档，命名一定要以.pdf结尾，不然在excel中打不开！！！。
          */
 
        // 创建画布和锚点
        Drawing<?> drawing = sheet.createDrawingPatriarch();
        // 将对象锚定在 "1 行 x 3 列" 的区域，初始就固定为该区域大小
        ClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, startCol, startRow, startCol + colsSpan, startRow + rowsSpan);
        // 固定大小/位置：单元格尺寸变化时，不随之移动或缩放
        anchor.setAnchorType(HSSFClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE);
        //文件略缩图会占据整个单元格，锚点随单元格大小的改变而自动调整。这意味着当单元格的大小发生变化时，图片的大小也会相应地进行调整。
        //在后续应用业务的时候，可以给单元格固定的高度和宽度，让略缩图更加美观
        //参数1：起始列的偏移量（单位为字符宽度的 1/256）
        //参数2：起始行的偏移量（单位为字符高度的 1/256）
        //参数3：结束列的偏移量（单位为字符宽度的 1/256）
        //参数4：结束行的偏移量（单位为字符高度的 1/256）
        //参数5：起始列
        //参数6：起始行
        //参数7：结束列
        //参数8：结束行
 
 
        // 创建图片并将它关联到OLE对象
        /**
         * 这里参数是：
         * 文件放在Excel表格内的位置：anchor
         * 文件在Excel表格内的索引：pdfIdx（文件本身，文件能否打开的关键）
         * 文件在Excel表格内的图标：iconid（文件的图标）
         * */
          drawing.createObjectData(anchor, oleIdx, iconid);
 
 
		//输出到 .xls 文件（若原为 .xlsx，则输出到同名 .xls）
		try (OutputStream outputStream = new FileOutputStream(outputPath)) {
			workbook.write(outputStream);
			System.out.println("文件写入成功：" + outputPath);
			if (lower.endsWith(".xlsx")) {
				System.out.println("提示：.xlsx 无法直接嵌入 OLE，已在同目录生成/更新 .xls 文件。");
			}
		}
        workbook.close();
    }
 
 
 
    // 将 Icon 对象转换为 Image 对象
    private static Image iconToImage(Icon icon) {
        if (icon instanceof ImageIcon) {
            return ((ImageIcon) icon).getImage();
        } else {
            int width = icon.getIconWidth();
            int height = icon.getIconHeight();
            BufferedImage image = new BufferedImage(width, height, BufferedImage.TYPE_INT_ARGB);
            Graphics graphics = image.createGraphics();
            icon.paintIcon(null, graphics, 0, 0);
            graphics.dispose();
            return image;
        }
    }
 
    // 将 Image 对象转换为字节数组
    private static byte[] imageToByteArray(Image image) throws IOException {
        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        ImageIO.write(imageToBufferedImage(image), "png", byteArrayOutputStream);
        return byteArrayOutputStream.toByteArray();
    }

    // 重载，按固定宽高缩放图像
    private static byte[] imageToByteArray(Image image, int width, int height) throws IOException {
        Image scaled = image.getScaledInstance(width, height, Image.SCALE_SMOOTH);
        BufferedImage bufferedImage = new BufferedImage(width, height, BufferedImage.TYPE_INT_ARGB);
        Graphics2D g2d = bufferedImage.createGraphics();
        g2d.setComposite(AlphaComposite.Src);
        g2d.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BILINEAR);
        g2d.setRenderingHint(RenderingHints.KEY_RENDERING, RenderingHints.VALUE_RENDER_QUALITY);
        g2d.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
        g2d.drawImage(scaled, 0, 0, null);
        g2d.dispose();

        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        ImageIO.write(bufferedImage, "png", byteArrayOutputStream);
        return byteArrayOutputStream.toByteArray();
    }
 
    // 将 Image 对象转换为 BufferedImage 对象
    private static BufferedImage imageToBufferedImage(Image image) {
        if (image instanceof BufferedImage) {
            return (BufferedImage) image;
        }
 
        BufferedImage bufferedImage = new BufferedImage(image.getWidth(null), image.getHeight(null), BufferedImage.TYPE_INT_ARGB);
        Graphics2D graphics2D = bufferedImage.createGraphics();
        graphics2D.drawImage(image, 0, 0, null);
        graphics2D.dispose();
        return bufferedImage;
    }
 
    // 将像素转换为列宽单位（1/256 字符宽度），近似换算
    private static int pixelsToColumnWidthUnits(int pixels) {
        double excelUnits = (pixels / 7.0 + 0.5) * 256.0;
        return (int) Math.round(excelUnits);
    }

    // 将像素转换为磅值（points），假设 96 DPI：1 px ≈ 0.75 pt
    private static float pixelsToPoints(int pixels) {
        return (float) (pixels * 72.0 / 96.0);
    }

}