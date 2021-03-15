import edu.stanford.nlp.ling.CoreAnnotations;
import edu.stanford.nlp.ling.CoreLabel;
import edu.stanford.nlp.pipeline.Annotation;
import edu.stanford.nlp.pipeline.StanfordCoreNLP;
import edu.stanford.nlp.util.CoreMap;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.*;
import java.util.*;

public class NLP {
    static XSSFRow row;
    public static void main(String[] args) throws IOException {
        ArrayList<Person> people = new ArrayList<>();
        try {
            readFromExcel("People.xlsx", people);
        } catch (FileNotFoundException e) {
            System.out.println("Ошибка! Файл Peoples.xlsx не найден.");
            System.in.read();
            return;
        }
        File file = new File("output.txt");
        FileWriter writer = new FileWriter(file);
        XWPFRun paragraphConfig;
        XWPFDocument docxModel;
        try {
            docxModel = new XWPFDocument();
            XWPFParagraph bodyParagraph = docxModel.createParagraph();
            bodyParagraph.setAlignment(ParagraphAlignment.THAI_DISTRIBUTE);
            paragraphConfig = bodyParagraph.createRun();
            paragraphConfig.setFontSize(14);
        } catch (Exception e) {
            e.printStackTrace();
            return;
        }

        String text;
        try {
            text = readText("input.txt");
        } catch (FileNotFoundException e) {
            System.out.println("Ошибка! Файл input.txt не найден.");
            System.in.read();
            return;
        }

        System.out.println("NLP is running...");

        String tag = null;
        boolean isTaged = false, isFind = false;

        Properties props = new Properties();
        props.setProperty("annotators", "tokenize, ssplit, pos, lemma, ner, parse, dcoref");
        StanfordCoreNLP pipeline = new StanfordCoreNLP(props);

        Annotation document = new Annotation(text);
        pipeline.annotate(document);

        List<CoreMap> sentences = document.get(CoreAnnotations.SentencesAnnotation.class);
        for (CoreMap sentence : sentences) {
            for (CoreLabel token : sentence.get(CoreAnnotations.TokensAnnotation.class)) {
                String word = token.get(CoreAnnotations.TextAnnotation.class);
                //String pos = token.get(CoreAnnotations.PartOfSpeechAnnotation.class);
                String ne = token.get(CoreAnnotations.NamedEntityTagAnnotation.class);
                String lemma = token.get(CoreAnnotations.LemmaAnnotation.class);

                if (!(ne.equals("O") || isTaged)) {
                    for (Person person : people) {
                        if (person.surname.contains(lemma) || person.name.contains(lemma) || person.patronymic.contains(lemma)
                                || (person.birthday.contains(lemma) && lemma.length() > 2)
                                || (person.address.contains(lemma) && lemma.length() > 4)
                                || (person.number.contains(lemma) && lemma.length() > 6)) {
                            isFind = true;
                            tag = person.guid;
                            break;
                        }
                    }
                    if (isFind) {
                        writer.write(" <" + tag + ">");
                        paragraphConfig.setText(" <" + tag + ">");
                        writer.flush();
                        //System.out.println(" <" + tag + ">");
                        isTaged = true;
                        //isFind = false;
                    }
                }
                if (ne.equals("O") && isTaged) {
                    writer.write(" </" + tag + ">");
                    paragraphConfig.setText(" </" + tag + ">");
                    writer.flush();
                    //System.out.println(" </" + tag + ">");
                    isTaged = false;
                }
                if (".,;!?:n't)".contains(word)) {
                    writer.write(word);
                    paragraphConfig.setText(word);
                }
                else {
                    writer.write(" " + word);
                    paragraphConfig.setText(" " + word);
                }
                writer.flush();
                //System.out.println(String.format("Print: word: [%s] pos: [%s] ne: [%s] lem: [%s]", word, pos, ne, lemma));
            }
            if (isTaged) {
                writer.write(" </" + tag + ">");
                paragraphConfig.setText(" </" + tag + ">");
            }
            writer.write("\n");
        }
        writer.close();
        FileOutputStream outputStream = new FileOutputStream("output.docx");
        docxModel.write(outputStream);
        outputStream.close();
        System.out.println("\nThe process has completed successfully.\nProcessed text located in files \"output.txt\" and \"output.docx\".\nPress Enter");
        System.in.read();
    }

    public static void readFromExcel(String file, ArrayList<Person> arr) throws IOException {
        FileInputStream fin = new FileInputStream(file);
        XSSFWorkbook wb = new XSSFWorkbook(fin);
        XSSFSheet spreadsheet = wb.getSheetAt(0);
        Iterator < Row >  rowIterator = spreadsheet.iterator();

        for (int i = 0; rowIterator.hasNext(); ++i) {
            row = (XSSFRow) rowIterator.next();
            Iterator < Cell >  cellIterator = row.cellIterator();
            arr.add(new Person());
            for (int j = 0; cellIterator.hasNext(); ++j) {
                Cell cell = cellIterator.next();
                if (i > 0) {
                    switch (j) {
                        case 0: arr.get(i - 1).guid = cell.getStringCellValue(); break;
                        case 1: arr.get(i - 1).surname = cell.getStringCellValue(); break;
                        case 2: arr.get(i - 1).name = cell.getStringCellValue(); break;
                        case 3: arr.get(i - 1).patronymic = cell.getStringCellValue(); break;
                        case 4: arr.get(i - 1).birthday = cell.getStringCellValue(); break;
                        case 5: arr.get(i - 1).address = cell.getStringCellValue(); break;
                        case 6: arr.get(i - 1).number = cell.getStringCellValue(); break;
                        default: j = 0;
                    }
                } else cell.getStringCellValue();
            }
        }
        fin.close();
    }

    public static String readText(String filename) throws IOException {
        try(BufferedReader br = new BufferedReader(new FileReader(filename))) {
            StringBuilder sb = new StringBuilder();
            String line = br.readLine();

            while (line != null) {
                sb.append(line);
                sb.append(System.lineSeparator());
                line = br.readLine();
            }
            return sb.toString();
        }
    }
}

