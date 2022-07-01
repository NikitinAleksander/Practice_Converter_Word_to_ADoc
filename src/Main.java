import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import java.io.*;
import java.lang.String;
import java.math.BigInteger;
import java.nio.charset.StandardCharsets;
import java.util.Iterator;
import java.util.List;

public class Main {
    public static void main(String[] args) {
        try (FileInputStream fileInputStream = new FileInputStream("C:\\Users\\Александр\\IdeaProjects\\Word_to_ADoc\\data\\a.docx")) {
            File file = new File("C:\\Users\\Александр\\IdeaProjects\\Word_to_ADoc\\data\\a.adoc");
            File fileLog = new File("C:\\Users\\Александр\\IdeaProjects\\Word_to_ADoc\\data\\a.log");
            Writer writer = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(file), StandardCharsets.UTF_8));
            Writer writerLog = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(fileLog), StandardCharsets.UTF_8));
            XWPFDocument docxFile = new XWPFDocument(OPCPackage.open(fileInputStream));

            String style = "";
            Iterator<IBodyElement> iter = docxFile.getBodyElementsIterator();
            String lastParagraph="";
            boolean endSection=true;
            boolean firstParagraph=true;
            int paragraph=1;
            int Table=1;
            int drawing=1;
            int picture=1;

            while (iter.hasNext()) {
                IBodyElement elem = iter.next();

                if (elem instanceof XWPFParagraph) {

                    //получаем параграф
                    XWPFParagraph p = (XWPFParagraph) elem;
                    String ParagraphString = "";
                    style = "";
                    boolean changeLast=false;

                    if (p.getRuns().size() != 0) {
                        //Это рисунок?
                        if (p.getRuns().get(0).getCTR().getDrawingArray().length == 1) {
                            writer.write("\n- Drawing"+drawing+"\n");
                            writerLog.write("Paragraph "+paragraph+" : Drawing "+drawing+" not migrate\n");
                            paragraph++;
                            drawing++;
                            lastParagraph="Drawing";
                            continue;
                        }
                        //Это картинка?
                        if (p.getRuns().get(0).getCTR().getPictArray().length == 1) {
                            writer.write("\n- Image"+picture+"\n");
                            writerLog.write("Paragraph "+paragraph+" : Picture "+picture+" not migrate\n");
                            paragraph++;
                            picture++;
                            lastParagraph="Image";
                            continue;
                        }
                        //Если мы здесь то это просто текст
                        //Просматриваем текст и его форматирование
                        for (XWPFRun run : p.getRuns()) {
                            String StyleText = "";
                            if (run.isBold()) {
                                StyleText+= "**";
                            }
                            if (run.isItalic()) {
                                StyleText+= "__";
                            }
                            ParagraphString += StyleText + run.text() + StyleText;
                        }
                    }
                    //Текст готов, узнаем стиль текста
                    if (p.getStyle() == null) {
                        style = "";
                    } else {
                        //Заголовок
                        if (p.getStyle().length() == 1 || p.getStyle().contains("Heading") || p.getStyle().contains("a3") || p.getStyle().contains("a5")) {
                            style = "=";
                            String pStyle = p.getStyle();
                            int HeadingLevel=0;
                            if (pStyle.contains("a3")) {
                                HeadingLevel = 1;
                            } else if (pStyle.contains("a5")) {
                                HeadingLevel = 2;
                            } else {
                                HeadingLevel = Integer.parseInt(pStyle.replace("Heading", ""));
                                if (HeadingLevel > 5) {
                                    HeadingLevel = 5;
                                }
                            }
                            for (int i = 0; i < HeadingLevel; i++) {
                                style += "=";
                            }
                            style += " ";
                            changeLast=true;
                            lastParagraph="Heading";
                        } else if (p.getNumID() != null) {
                            //список
                            BigInteger m = new BigInteger(Integer.toString(6));
                            style = "*";
                            if (p.getNumID() != null && p.getNumID().compareTo(m) == 0) {
                                style = ".";
                            }
                            if (p.getNumIlvl() != null) {
                                for (int i = 0; i < p.getNumIlvl().intValue(); i++) {
                                    if (style.contains(".")) {
                                        style += ".";
                                    } else {
                                        style += "*";
                                    }
                                }
                            }
                            style += " ";
                            if(lastParagraph.equals("List")){
                                endSection=false;
                            }
                            changeLast=true;
                            lastParagraph="List";
                        }else {
                            lastParagraph="";
                            changeLast=true;
                            if (p.getStyle().equals("Title")){
                                style="= ";
                            }
                        }

                    }
                    if (!style.contains("=")){
                        if (lastParagraph.equals("Heading")){
                            endSection=false;
                        }
                    }
                    if (p.getText()==""){
                        continue;
                    }
                    if(!firstParagraph){
                        if(endSection){
                            writer.write("\n");
                        }else {
                            endSection=true;
                        }
                    }else{
                        firstParagraph=false;
                    }
                    writer.write(style +ParagraphString);
                    writer.write('\n');
                    writerLog.write("Paragraph "+paragraph);
                    if(p.getText()!=""){
                        writerLog.write(" \""+p.getRuns().get(0).text()+"\"");
                    }
                    writerLog.write(" : Paragraph migrate\n");
                    paragraph++;

                    if(!changeLast){
                        lastParagraph="";
                    }
                } else if (elem instanceof XWPFTable) {
                    XWPFTable table = (XWPFTable) elem;
                    writer.write("\n" + "|==="+"\n");
                    List<XWPFTableRow> rows = table.getRows();
                    List<XWPFTableCell> cells;
                    String row = "";
                    for (int i = 0; i < rows.size(); i++) {
                        cells = rows.get(i).getTableCells();
                        for (int j = 0; j < cells.size(); j++) {
                            int tmp=1;
                            int VMerge=1;
                            BigInteger HMerge= new BigInteger(Integer.toString(1));
                            if (rows.get(i).getCell(j).getCTTc().getTcPr().getVMerge() != null){
                                while((i+tmp< rows.size())&&(rows.get(i+tmp).getCell(j).getCTTc().getTcPr().getVMerge()!= null)&&(rows.get(i+tmp).getCell(j).getCTTc().getTcPr().getVMerge().getVal()== null)){
                                    VMerge++;
                                    tmp++;
                                }
                            }
                            if (rows.get(i).getCell(j).getCTTc().getTcPr().getGridSpan()!= null){
                                if(rows.get(i).getCell(j).getCTTc().getTcPr().getVMerge()==null){
                                    HMerge = rows.get(i).getCell(j).getCTTc().getTcPr().getGridSpan().getVal();
                                }else if(rows.get(i).getCell(j).getCTTc().getTcPr().getVMerge().getVal()!=null){
                                    HMerge = rows.get(i).getCell(j).getCTTc().getTcPr().getGridSpan().getVal();
                                }else{
                                    continue;
                                }
                            }
                            if(VMerge>1||HMerge.intValue()>1){
                                writer.write(Integer.toString(HMerge.intValue())+"."+Integer.toString(VMerge)+"+|"+cells.get(j).getText()+" ");
                            }else {
                                if (HMerge.intValue()==1&&(rows.get(i).getCell(j).getCTTc().getTcPr().getVMerge()!=null)&&rows.get(i).getCell(j).getCTTc().getTcPr().getVMerge().getVal()==null){
                                    continue;
                                }
                                writer.write("|" + cells.get(j).getText()+" ");
                            }
                        }
                        writer.write("\n");
                    }
                    writer.write("|===" + "\n");
                    writerLog.write("Paragraph "+paragraph+" : Table "+Table+" migrate\n");
                    paragraph++;
                    Table++;
                }
            }
            writer.close();
            writerLog.close();
        } catch (
                Exception ex) {
            ex.printStackTrace();
        }
    }
}


