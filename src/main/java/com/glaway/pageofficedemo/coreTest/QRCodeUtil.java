package com.glaway.pageofficedemo.coreTest;

import com.google.zxing.BarcodeFormat;
import com.google.zxing.EncodeHintType;
import com.google.zxing.MultiFormatWriter;
import com.google.zxing.WriterException;
import com.google.zxing.common.BitMatrix;
import com.google.zxing.qrcode.decoder.ErrorCorrectionLevel;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.geom.RoundRectangle2D;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.IOException;
import java.util.Hashtable;
import java.util.UUID;

/**
 * 二维码生成类
 * @author guiqingle
 * @version 0.1
 *
 */
public class QRCodeUtil {

    private static final String CHARSET = "utf-8";
    private static final String FORMAT_NAME = "jpg";
    //二维码尺寸
    private static final int QRCODE_SIZE = 300;
    //logo宽度
    private static final int WIDTH = 60;
    //logo高度
    private static final int HEIGHT = 60;

    private static BufferedImage createImage(String content,String imgPath,boolean needConpress) throws Exception {

        Hashtable<EncodeHintType,Object> hints = new Hashtable<EncodeHintType,Object>();
        hints.put(EncodeHintType.ERROR_CORRECTION, ErrorCorrectionLevel.H);
        hints.put(EncodeHintType.CHARACTER_SET,CHARSET);
        hints.put(EncodeHintType.MARGIN,1);
        BitMatrix bitMatrix = new MultiFormatWriter().encode(content, BarcodeFormat.QR_CODE,QRCODE_SIZE,QRCODE_SIZE,hints);

        int width = bitMatrix.getWidth();
        int height = bitMatrix.getHeight();
        BufferedImage image = new BufferedImage(width,height,BufferedImage.TYPE_INT_RGB);
        for (int x = 0; x < width;x++){
            for (int y = 0; y < height; y++){
                image.setRGB(x,y,bitMatrix.get(x,y) ? 0xFF000000 : 0xFFFFFFFF);
            }
        }
        if (imgPath == null || "".equals(imgPath)){
            return image;
        }
        //插入图片
        QRCodeUtil.insertImage(image,imgPath,needConpress);
        return image;

    }

    /**
     * 插入LOGO
     * @param source 二维码图片
     * @param imgPath LOGO图片地址
     * @param needConpress 是否压缩
     */
    private static void insertImage(BufferedImage source, String imgPath, boolean needConpress) throws Exception {
        File file = new File(imgPath);
        if (!file.exists()){
            System.err.println("" + imgPath + "，该文件不存在！");
            return;
        }
        Image src = ImageIO.read(file);
        int width = src.getWidth(null);
        int height = src.getHeight(null);

        if(needConpress){ //压缩logo
            if (width > WIDTH){
                width = WIDTH;
            }
            if (height > HEIGHT){
                height = HEIGHT;
            }
            Image image = src.getScaledInstance(width,height,Image.SCALE_SMOOTH);
            BufferedImage tag = new BufferedImage(width,height,BufferedImage.TYPE_INT_RGB);

            Graphics g = tag.getGraphics();
            g.drawImage(image,0,0,null);//绘制缩小后的图
            g.dispose();
            src = image;
        }

        //插入LOGO
        Graphics2D graph = source.createGraphics();
        int x = (QRCODE_SIZE -width) / 2;
        int y = (QRCODE_SIZE - height) / 2;
        graph.drawImage(src,x,y,width,height,null);
        Shape shape = new RoundRectangle2D.Float(x,y,width,height,6,6);
        graph.setStroke(new BasicStroke(3f));
        graph.draw(shape);
        graph.dispose();
    }

    /**
     * 生成二维码（内嵌LOGO）
     * @param content 内容
     * @param imgPath LOGO地址
     * @param destPath 存放位置
     * @param needCompress 是否压缩LOGo
     * @return
     * @throws Exception
     */
    public  static  String encode(String content,String imgPath,String destPath,boolean needCompress) throws Exception{
        BufferedImage image = QRCodeUtil.createImage(content,imgPath,needCompress);
        mkdirs(destPath);
        //随机生成二维码图片文件名
        String file  = UUID.randomUUID() + ".jpg";
        ImageIO.write(image,FORMAT_NAME,new File(destPath + "/" + file));
        return destPath + file;
    }

    /**
     * 当文件夹不存在时，mkdirs会自动创建多层目录，区别于mkdir．(mkdir如果父目录不存在则会抛出异常)
     *
     * @author lanyuan Email: mmm333zzz520@163.com
     * @date 2013-12-11 上午10:16:36
     * @param destPath 存放目录
     */
    public static void mkdirs(String destPath) {
        File file = new File(destPath);
        // 当文件夹不存在时，mkdirs会自动创建多层目录，区别于mkdir．(mkdir如果父目录不存在则会抛出异常)
        if (!file.exists() && !file.isDirectory()) {
            file.mkdirs();
        }
    }


    public static void main(String[] args) throws Exception {
        String content = "这是测试二维码的内容";
        String imgPath = "C:\\Users\\lenovo\\Pictures\\Saved Pictures\\lo.jpg";
        String destPath = "C:\\Users\\lenovo\\Pictures\\Saved Pictures";
        new QRCodeUtil().encode(content,imgPath,destPath,true);
    }
}
