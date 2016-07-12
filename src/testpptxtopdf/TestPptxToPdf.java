/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package testpptxtopdf;

import java.awt.Color;
import java.awt.Dimension;
import java.awt.Graphics2D;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
/**
 *
 * @author sagar
 */
public class TestPptxToPdf {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws Exception {
        // TODO code application logic here
        pptxToPdf();
    }
    
    static void pptxToPdf() throws Exception{
        String dir = "/home/sagar/Desktop/Shareback/";
        String filepath = dir+"test.pptx";
        
        //creating an empty presentation
      File file=new File(filepath);
      XMLSlideShow ppt = new XMLSlideShow(new FileInputStream(file));
      
      //getting the dimensions and size of the slide 
        Dimension pgsize = ppt.getPageSize();
      XSLFSlide[] slide = ppt.getSlides();
      
      int idx=1;
      for (int i = 0; i < slide.length; i++) {
          BufferedImage img = new BufferedImage(pgsize.width, pgsize.height,BufferedImage.TYPE_INT_RGB);
          Graphics2D graphics = img.createGraphics();

         //clear the drawing area
         graphics.setPaint(Color.white);
         graphics.fill(new Rectangle2D.Float(0, 0, pgsize.width, pgsize.height));

         //render
         slide[i].draw(graphics);
         
         //creating an image file as output
         
         File f = new File(dir+"ptx/");
         if(!f.exists()){
             f.mkdir();
         }
         FileOutputStream out = new FileOutputStream(f.getAbsolutePath()+"img"+idx+".png");
         
         javax.imageio.ImageIO.write(img, "png", out);
         ppt.write(out);

         System.out.println("Image successfully created");
         out.close();
         idx++;
      }	
    }
    
}
