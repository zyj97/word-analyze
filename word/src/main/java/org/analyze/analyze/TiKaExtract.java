package org.analyze.analyze;

import javafx.scene.effect.SepiaTone;
import org.apache.tika.exception.TikaException;
import org.apache.tika.metadata.Metadata;
import org.apache.tika.mime.MediaType;
import org.apache.tika.parser.ParseContext;
import org.apache.tika.parser.Parser;
import org.apache.tika.parser.microsoft.OfficeParser;
import org.apache.tika.sax.BodyContentHandler;
import org.springframework.stereotype.Component;
import org.springframework.web.multipart.MultipartFile;
import org.w3c.dom.Document;
import org.xml.sax.SAXException;

import java.io.IOException;
import java.io.InputStream;
import java.util.Set;

import static org.apache.tika.parser.microsoft.OfficeParser.POIFSDocumentType.WORDDOCUMENT;

/**
 * @Author: zyj
 * @Date: 2021/6/7
 *
 * Apache Tika
 */
@Component
public class TiKaExtract {
    public void extract(MultipartFile file) throws IOException, TikaException, SAXException {
        BodyContentHandler handler = new BodyContentHandler(1024 * 1024 * 10);
        Metadata metadata = new Metadata();
        // Tika-1.1最高支持2007及更低版本的Office Word文档，如果是高于2007版本的Word文档需要使用POI处理（Tika会报错）
        InputStream inputstream = file.getInputStream();
        ParseContext pcontext = new ParseContext();

        // 解析Word文档时应由超类AbstractParser的派生类OfficeParser实现
        Parser msofficeparser = new OfficeParser();
        msofficeparser.parse(inputstream, handler, metadata, pcontext);
        Document document = pcontext.getDocumentBuilder().newDocument();
        // 获取Word文档的内容
        Set<MediaType> a = msofficeparser.getSupportedTypes(pcontext);
        MediaType mediaType = a.iterator().next();


        System.out.println("Word文档内容:" + handler.toString());

        // 获取Word文档的元数据
        System.out.println("Word文档元数据:");
        String[] metadataNames = metadata.names();

        for (String name : metadataNames) {
            System.out.println(name + " : " + metadata.get(name));
        }
    }
}






