package com.yw.easy_excel_test.utils;


import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFObjectData;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.IOUtils;

import javax.imageio.ImageIO;
import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.swing.*;
import javax.swing.filechooser.FileSystemView;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.awt.image.RenderedImage;
import java.io.*;

public class CSPDF  {



    public static void main(String[] args) throws IOException {
        //只有HSSFWorkbook才能使用OLE对象，并且poi需要在4.0之上，相关jar可以在我的资源中下载，不需要积分，也不需要VIP！！！
        // 创建工作簿和工作表对象
        Workbook workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet("Sheet1");
        sheet.setDefaultColumnWidth((short) 20);
        // 固定第一行第三列（行索引0，列索引2）的单元格尺寸
        sheet.setColumnWidth(2, 20 * 256);
        Row row0 = sheet.getRow(0) != null ? sheet.getRow(0) : sheet.createRow(0);
        row0.setHeightInPoints(80F);
        // 读取需要添加的PDF文件，如果有需要添加多个文件的需求，可以循环表格来添加
        File pdfFile = new File("D:\\Dsesktop\\test.docx");
        FileInputStream fis = new FileInputStream(pdfFile);
        byte[] pdfBytes = new byte[(int) pdfFile.length()];
        fis.read(pdfBytes);
        fis.close();
//        // 获取文件系统视图
//        FileSystemView view = FileSystemView.getFileSystemView();
//        // 获取PDF文件的展示图标
//        String imagePath = "D:/123456.png"; // 图片文件路径 这个是excel中单元格中pdf文件对应展示出来的图片，可以自定义
//        BufferedImage image = ImageIO.read(new File(imagePath));
//        ByteArrayOutputStream baos = new ByteArrayOutputStream();
//        ImageIO.write(image, "png", baos);
//        byte[] imageBytes = baos.toByteArray();
//        baos.close();
//        /** 这一步是获取文件图标 */
//        // 获取文件系统视图
//        FileSystemView view = FileSystemView.getFileSystemView();
//        Icon systemIcon = view.getSystemIcon(pdfFile);
//        // 将 Icon 对象转换为 Image 对象
//        Image image = iconToImage(systemIcon);
//        // 将 Image 对象转换为字节数组
//        byte[] imageBytes = imageToByteArray(image);

        /**  自定义文件的展示图标 */
        String imagePath = "D:\\xfg\\easy_excel_test\\src\\main\\java\\com\\yw\\easy_excel_test\\utils\\img.png"; // 图片文件路径 这个是excel中单元格中pdf文件对应展示出来的图片，可以自定义
        BufferedImage image = ImageIO.read(new File(imagePath));
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        ImageIO.write(image, "png", baos);
        byte[] imageBytes = baos.toByteArray();
        baos.close();

        int iconid = workbook.addPicture(imageBytes, HSSFWorkbook.PICTURE_TYPE_PNG);//将图片添加进入到Excel文件内
        int pdfIdx = workbook.addOlePackage(pdfBytes, "test.docx", "test.docx", "test.docx");// 在工作表中创建OLE对象
//        pdfBytes: 表示 PDF 文件的字节数组。您需要将 PDF 文件的内容以字节数组的形式传递给此参数。
//        "20230705.pdf": 表示 OLE 对象的类型。如果导出之后，在excel中操作另存文件，这个名称就是新保存文件的默认名称
//        "D:/file8261.pdf": 这指定了要在OLE Package中使用的类名。在此示例中，你可以将其设置为任何你想要的类名，因为它不会对后续的操作产生影响。它也只是用于标识特定的OLE Package。
//        "cs.pdf":这是在Excel工作簿中显示的OLE Package的名称,如果需要在excel中打开pdf文档，命名一定要以.pdf结尾，不然在excel中打不开！！！。

        // 创建画布和锚点
        CreationHelper helper = workbook.getCreationHelper();
        Drawing<?> drawing = sheet.createDrawingPatriarch();
        ClientAnchor anchor = drawing.createAnchor(0, 0, 100, 50, 2, 0, 3, 1);// 固定在第一行第三列，铺满该单元格
        anchor.setAnchorType(HSSFClientAnchor.AnchorType.MOVE_AND_RESIZE);
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
        HSSFObjectData objectData = (HSSFObjectData) drawing.createObjectData(anchor, pdfIdx, iconid);//设置缩略图和文件锚点的关系
        // 保存工作簿至文件
        //JAVA语言版本最低需要是7，否则会报错
        try (OutputStream outputStream = new FileOutputStream("D:\\xfg\\easy_excel_test\\src\\main\\java\\com\\yw\\easy_excel_test\\utils\\CS.xlsx")) {//excel保存的路径是自定义的，可以修改成任意路径
            workbook.write(outputStream);
        }
        workbook.close();
    }

    private static byte[] readFileToByteArray(File file) throws IOException {
        try (FileInputStream fis = new FileInputStream(file)) {
            return IOUtils.toByteArray(fis);
        }
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
}