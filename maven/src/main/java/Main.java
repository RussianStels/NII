
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.xmlbeans.XmlException;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Created by mayro on 13.02.2016.
 */

public class Main {

    public static void main(String[] args) throws IOException, OpenXML4JException, XmlException {
        File file = new File("C:\\Users\\mayro\\Desktop\\NII\\maven\\src\\main\\resources\\med_map.doc");
        File resFile = new File("C:\\Users\\mayro\\Desktop\\NII\\maven\\target\\result.doc");

        POIFSFileSystem poifsFileSystem = new POIFSFileSystem(file);
        HWPFDocument document = new HWPFDocument(poifsFileSystem);
        Range range = document.getRange();
        range.replaceText("[MED_CARD_NUM]","123");
        range.replaceText("[pat_card]","666");
        range.replaceText("[DATE_POST]","12:02");
        range.replaceText("[FIO]","Пупкин Вася Батькович");
        range.replaceText("[postupl]","Экстренный");
        range.replaceText("[diagpost]","Болен. Очень болен. ");
        range.replaceText("[stardoc]","Быков :D");
        range.replaceText("[SEX]","M");
        range.replaceText("[age]","101");
        range.replaceText("[DATE_POST]","10.10.2010");
        range.replaceText("[DATE_OUT]","10.10.2010");
        range.replaceText("[DEP_NAME]","Терапевтия или как там");
        range.replaceText("[BED]","1");
        range.replaceText("[DEP_NAME_OUT]","Хирургия");
        range.replaceText("[KD]","21");
        range.replaceText("[rh_blood_code]","4+");
        FileOutputStream fileOutputStream = new FileOutputStream(resFile);
        document.write(fileOutputStream);

        WordExtractor wordExtractor = new WordExtractor(poifsFileSystem);
        System.out.println(wordExtractor.getText());
        fileOutputStream.close();
        poifsFileSystem.close();
//        file.delete();
//        resFile.renameTo(file);
    }
}
