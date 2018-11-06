package com.workiva.asposetest;

import com.aspose.slides.*;

import java.io.File;
import java.nio.file.Files;

public class Main {
    public static void main(String[] args) {
        String dir = System.getProperty("user.dir");

        Presentation presentation = new Presentation(dir + "/asposetest.pptx");

        ISlide currentSlide = presentation.getSlides().get_Item(0);

        IShape shape = currentSlide.getShapes().get_Item(0);

        IPictureFrame picture = null;
        if(shape instanceof IPictureFrame) {
            picture = (IPictureFrame)shape;
        }

        IShapeStyle style = picture.getShapeStyle(); // style is null

        if(style == null) {
            System.out.println("style is null");
        }

        File file = new File("test.png");
        try {
            byte[] fileContent = Files.readAllBytes(file.toPath());

            IPPImage image = presentation.getImages().addImage(fileContent);

            IPictureFrame newPicture = currentSlide.getShapes().addPictureFrame(ShapeType.Rectangle, 0, 0, 100, 100, image);

            IShapeStyle newStyle = newPicture.getShapeStyle(); // always returns null

            if(newStyle == null) {
                System.out.println("newStyle is null");
            }
        } catch(Exception ex) {
            // ignore exception for now
        }
    }
}
