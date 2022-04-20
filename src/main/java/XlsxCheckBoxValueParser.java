import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.*;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.xml.sax.*;
import org.xml.sax.helpers.DefaultHandler;

import javax.xml.parsers.ParserConfigurationException;
import java.io.IOException;
import java.net.URI;
import java.util.*;
import java.util.stream.Collectors;

/**
 *  一个基于 Apache POI 提供的 SAX 解析，来解析 xlsx 中 checkBox 的的小工具
 *
 */
public class XlsxCheckBoxValueParser {

    private final ExtXSSReader extXssReader;  // 基于 POI 提供的 XSSReader 扩展了几个需要的 API
    private final SheetChooseHandler sheetChooseHandler; // 维护选中 sheet 信息
    private SheetCheckBoxHandler sheetCheckBoxHandler;  // 处理选中 sheet 中的 checkBox

    private final XMLReader xmlReader;  // xml 解析器



    /**
     * 构造 checkBox 解析器时会初始化相应的处理器
     * @param xlsxFile 需要处理的 xlsx 文件路径
     * @throws OpenXML4JException
     * @throws ParserConfigurationException
     * @throws SAXException
     * @throws IOException
     */
    public XlsxCheckBoxValueParser(String xlsxFile) throws OpenXML4JException, ParserConfigurationException, SAXException, IOException {
        OPCPackage pkg = OPCPackage.open(xlsxFile, PackageAccess.READ);

        extXssReader = new ExtXSSReader(pkg);

        xmlReader = XMLHelper.newXMLReader();


        // sheetChooseHandler 根据 workbook.xml 中的信息建立 sheet name 到 sheet rid 的映射
        // 并维护选中的 sheet 信息
        // 通过 xmlReader 启动解析，把信息交给 sheetChooseHandler 处理
        sheetChooseHandler = new SheetChooseHandler();
        xmlReader.setContentHandler(sheetChooseHandler);
        xmlReader.parse(new InputSource(extXssReader.getWorkbookData()));


    }

    /**
     * 选择一个 sheet 读取
     * 维护 sheetChooseHandler 的选择信息
     * 并通过 SheetCheckBoxHandler 解析 sheet 中的 checkBox
     * @param sheetName 选中 sheet 的名字
     * @throws IOException
     * @throws InvalidFormatException
     * @throws SAXException
     * @throws ParserConfigurationException
     */
    public void chooseSheet(String sheetName) throws IOException, InvalidFormatException, SAXException, ParserConfigurationException {
        if (sheetChooseHandler != null) {

            sheetChooseHandler.setChosenSheetName(sheetName);



            // sheetCheckBoxHandler 会处理 worksheets/sheet#.xml 中的 checkBox 信息并构造 CheckBox 对象
            // sheet#.xml 由 sheetChooseHandler 维护的选中 sheet 确定
            // 同样通过 xmlReader 启动解析后通过 sheetCheckBoxHandler 处理
            PackagePart chosenSheetPart = extXssReader.getSheetPartById(sheetChooseHandler.getChosenSheetRid());
            sheetCheckBoxHandler = new SheetCheckBoxHandler(sheetName,
                    sheetChooseHandler.getChosenSheetRid(), chosenSheetPart);
            xmlReader.setContentHandler(sheetCheckBoxHandler);
            xmlReader.parse(new InputSource(extXssReader.getSheet(sheetCheckBoxHandler.getSheetRid())));

        }
    }

    /**
     * 读取一个 CheckBox 的选择值
     * @param row checkBox 所在的行
     * @param col checkBox 所在列
     * @param text checkBox 的文本内容
     * @return checkBox 是否选中
     * @throws InvalidFormatException
     * @throws IOException
     * @throws SAXException
     * @throws ParserConfigurationException
     */
    public boolean getCheckBoxValue(int row, int col, String text) throws InvalidFormatException, IOException, SAXException, ParserConfigurationException {
        boolean checked = false;

        XlsxCheckBox checkBox = getCheckBox(row, col, text);

        if (checkBox != null) {
            checked = checkBox.getCheckedValue();
        }

        return checked;
    }

    /**
     * 获取指定的 checkBox
     * @param row checkBox 所在行
     * @param col checkBox 所在列
     * @param text checkBox 的文本内容
     * @return
     */
    public XlsxCheckBox getCheckBox(int row, int col, String text) {
        XlsxCheckBox xlsxCheckBox = null;

        List<XlsxCheckBox> checkBoxes = this.sheetCheckBoxHandler.getXlsxCheckBox(row, col);

        for (XlsxCheckBox checkBox: checkBoxes) {
            if (Objects.equals(text, checkBox.getText())) {
                xlsxCheckBox = checkBox;
                break;
            }
        }

        return xlsxCheckBox;

    }

    /**
     * 获取指定行列(单元格)的 checkBox
     * 由于一个单元格内可能存在多个 checkBox, 因此返回 List
     * @param row
     * @param col
     * @return
     */
    public List<XlsxCheckBox> getCheckBox(int row, int col) {
        return this.sheetCheckBoxHandler.getXlsxCheckBox(row, col);
    }


    /**
     * 获取当前 sheet 下有 checkBox 的所有行索引
     * @return
     */
    public List<Integer> getCheckBoxRows() {
        return this.sheetCheckBoxHandler.getCheckBoxRows();
    }

    /**
     * 获取当前 sheet 下指定行下，有 checkBox 的所有列索引
     * @param row
     * @return
     */
    public List<Integer> getCheckBoxCols(int row) {
        return this.sheetCheckBoxHandler.getCheckBoxCols(row);
    }

    /**
     * 维护表格的 sheet 信息
     * 维护表格当前的 sheet 信息
     */
    private static class SheetChooseHandler extends DefaultHandler {

        private static final String SHEET = "sheet";
        private static final String RID = "id";
        private static final String NAME = "name";


        private final Map<String, String> nameRidMap = new HashMap<>();

        private String chosenSheetName = null;

        public String getChosenSheetRid() {
            return nameRidMap.get(chosenSheetName);
        }

        public void setChosenSheetName(String name) {
            this.chosenSheetName = name;
        }

        public String getChosenSheetName() {
            return this.chosenSheetName;
        }


        /**
         * 访问 workbook.xml 中的元素，根据格式样例解析目标元素
         * sheet 元素格式样例：
         *      <sheet name="Sheet6" sheetId="4" r:id="rId6"/>
         * @param uri
         * @param localName
         * @param qName
         * @param attrs
         * @throws SAXException
         */
        @Override
        public void startElement(String uri, String localName, String qName, Attributes attrs) throws SAXException {
            if (localName.equalsIgnoreCase(SHEET)) {
                String name = null;
                String rid = null;
                for (int i = 0; i < attrs.getLength(); i++) {
                    final String attrName = attrs.getLocalName(i);
                    if (attrName.equalsIgnoreCase(NAME)) {
                        name = attrs.getValue(i);
                    } else if (attrName.equalsIgnoreCase(RID)) {
                        rid = attrs.getValue(i);
                    }
                    if (name != null && rid != null) {
                        nameRidMap.put(name, rid);
                        break;
                    }
                }
            }
        }
    }

    /**
     * 解析 sheet 中的 checkBox
     * 维护当前 sheet 的 checkBox 信息
     */
    private static class SheetCheckBoxHandler extends DefaultHandler {

        private static final String CONTROL = "control";
        private static final String NAME = "name";
        private static final String RID = "id";
        private static final String FROM = "from";
        private static final String ROW = "row";
        private static final String COL = "col";

        private final String sheetName;
        private final String sheetRid;
        private final PackagePart chosenSheetPart;
        private final XMLReader drawingReader;

        public SheetCheckBoxHandler(String sheetName, String sheetRid, PackagePart chosenSheetPart) throws ParserConfigurationException, SAXException {
            this.sheetName = sheetName;
            this.sheetRid = sheetRid;
            this.chosenSheetPart = chosenSheetPart;

            this.drawingReader = XMLHelper.newXMLReader();
        }

        public String getSheetName() {
            return this.sheetName;
        }

        public String getSheetRid() {
            return this.sheetRid;
        }


        private final Map<Integer, Map<Integer, List<XlsxCheckBox>>> checkBoxMap = new HashMap<>();

        /**
         * 获取当前 sheet 下指定单元格的 checkBox
         * @param row
         * @param col
         * @return
         */
        public List<XlsxCheckBox> getXlsxCheckBox(int row, int col) {

            Map<Integer, List<XlsxCheckBox>> colCheckBoxMap = checkBoxMap.get(row);

            List<XlsxCheckBox> checkBoxes = null;

            if (colCheckBoxMap != null) {
                checkBoxes = colCheckBoxMap.getOrDefault(col, new ArrayList<>());
            }

            return checkBoxes;

        }


        public List<Integer> getCheckBoxRows() {
            return checkBoxMap.keySet().stream().sorted().collect(Collectors.toList());
        }

        public List<Integer> getCheckBoxCols(int row) {
            return checkBoxMap.getOrDefault(row, new HashMap<>()).keySet()
                    .stream().sorted().collect(Collectors.toList());
        }

        private XlsxCheckBox xlsxCheckBox;

        private Boolean inControl = false;
        private Boolean inFrom = false;
        private Boolean inCol = false;
        private Boolean inRow = false;

        /**
         * 解析 checkBox 的名字和 rid
         * @param attrs
         */
        public void parseCheckBoxNameAndRid(Attributes attrs) {
            xlsxCheckBox = new XlsxCheckBox();
            String name = null;
            String rid = null;
            for (int i = 0; i < attrs.getLength(); i++) {
                final String attrName = attrs.getLocalName(i);
                if (attrName.equalsIgnoreCase(NAME)) {
                    name = attrs.getValue(i);
                } else if (attrName.equalsIgnoreCase(RID)) {
                    rid = attrs.getValue(i);
                }
                if (name != null && rid != null) {
                    xlsxCheckBox.setName(name);
                    xlsxCheckBox.setRid(rid);
                    break;
                }
            }
        }

        /**
         * 解析 checkBox 的文本内容
         * 通过 DrawingCheckBoxTextContentHandler 来解析
         * @param xlsxCheckBox
         * @throws InvalidFormatException
         * @throws IOException
         * @throws SAXException
         */
        public void parseCheckBoxText(XlsxCheckBox xlsxCheckBox) throws InvalidFormatException, IOException, SAXException {

            // 构造 checkbox textContent 解析器
            DrawingCheckBoxTextContentHandler handler = new DrawingCheckBoxTextContentHandler(xlsxCheckBox);
            drawingReader.setContentHandler(handler);

            // 获取 sheet#.xml 对应的 drawing#.xml 文件
            // 文本内容在 drawing#.xml 文件中
            PackageRelationship drawingRelationship = chosenSheetPart.getRelationshipsByType(XSSFRelation.DRAWINGS.getRelation())
                    .getRelationship(0);
            PackagePart drawingPart = chosenSheetPart.getRelatedPart(drawingRelationship);

            // 通过 reader 启动解析
            drawingReader.parse(new InputSource(drawingPart.getInputStream()));

        }


        /**
         * 解析 checkBox 信息标签
         * 其元素标签样例如下：
         *      <control r:id="rId3" name="Check Box 14" shapeId="34830">
         *          ...
         *          <from>
         *              ...<xdr:col>checkbox 所在列</xdr:col> ...
         *              ...<xdr:row>checkbox 所在行</xdr:row>
         *          </from>
         *      </control>
         * @param uri
         * @param localName
         * @param qName
         * @param attrs
         * @throws SAXException
         */
        @Override
        public void startElement(String uri, String localName, String qName, Attributes attrs) throws SAXException {
            if (localName.equalsIgnoreCase(CONTROL)) {
                inControl = true;
                parseCheckBoxNameAndRid(attrs);
            }
            else if (inControl && localName.equalsIgnoreCase(FROM)) {
                inFrom = true;
            }
            else if(inFrom && localName.equalsIgnoreCase(ROW)) {
                inRow = true;
            }
            else if(inFrom && localName.equalsIgnoreCase(COL)) {
                inCol = true;
            }
        }

        /**
         * 根据标签 flag 解析相应的内容
         * @param ch
         * @param start
         * @param length
         * @throws SAXException
         */
        public void characters(char[] ch, int start, int length) throws SAXException {

            if (inCol) {
                // 当前标签是 <xdr:col>, 解析列索引
                String value = new String(ch, start, length);
                xlsxCheckBox.setColIndex(Integer.parseInt(value));
            }
            else if(inRow) {

                // 当前标签是 <xdr:row>, 解析行索引
                String value = new String(ch, start, length);
                xlsxCheckBox.setRowIndex(Integer.parseInt(value));
            }

        }

        /**
         * 关闭标签，处理一些 flag
         * 处理 checkBox 对象
         * @param uri
         * @param localName
         * @param name
         * @throws SAXException
         */
        public void endElement(String uri, String localName, String name) throws SAXException {

            if (inCol && localName.equalsIgnoreCase(COL)) {
                inCol = false;
            }

            if (inRow && localName.equalsIgnoreCase(ROW)) {
                inRow = false;
            }

            if (inFrom && localName.equalsIgnoreCase(FROM)) {
                inFrom = false;
            }

            if (inControl && localName.equalsIgnoreCase(CONTROL)) {
                inControl = false;

                try {
                    parseCheckBoxText(xlsxCheckBox);
                } catch (InvalidFormatException e) {
                    e.printStackTrace();
                } catch (IOException e) {
                    e.printStackTrace();
                }

                // workbook.xml.ref 记录有 sheet#.xml 的 ref 文件
                // 根据 sheet#.xml 的 ref 文件可以得到 checkBox 的 ctrlProp 文件
                PackageRelationship ctrlPropRelationship = chosenSheetPart.getRelationship(xlsxCheckBox.getRid());
                try {
                    PackagePart ctrlPropPart = chosenSheetPart.getRelatedPart(ctrlPropRelationship);
                    xlsxCheckBox.setCtrlPropPart(ctrlPropPart);
                } catch (InvalidFormatException e) {
                    e.printStackTrace();
                }

                // 按照 row 和 col 把 checkBox 维护在 Map 中
                int row = xlsxCheckBox.getRowIndex();
                Map<Integer, List<XlsxCheckBox>> checkBoxColMap = checkBoxMap.getOrDefault(row, new HashMap<>());
                List<XlsxCheckBox> checkBoxes = checkBoxColMap.getOrDefault(xlsxCheckBox.getColIndex(), new ArrayList<>());
                checkBoxes.add(xlsxCheckBox);
                checkBoxColMap.put(xlsxCheckBox.getColIndex(), checkBoxes);
                checkBoxMap.put(row, checkBoxColMap);





                xlsxCheckBox = null;
            }
        }

    }

    /**
     * 解析 checkbox 对应的 ctrlProp 信息来获取 checked value
     */
    private static class CtrlPropCheckBoxHandler extends DefaultHandler {
        private static final String FORM_CONTROL_PR = "formControlPr";
        private static final String CHECKED = "checked";

        private XlsxCheckBox xlsxCheckBox;
        private boolean checked = false;

        public CtrlPropCheckBoxHandler(XlsxCheckBox xlsxCheckBox) {
            this.xlsxCheckBox = xlsxCheckBox;
        }

        public boolean getChecked() {
            return this.checked;
        }

        /**
         * 解析 checkBox 的 checked 信息
         * 选中时标签格式为：
         *      <formControlPr val="0" noThreeD="1" checked="checked" objectType="CheckBox" xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"/>
         * @param uri
         * @param localName
         * @param qName
         * @param attrs
         * @throws SAXException
         */
        @Override
        public void startElement(String uri, String localName, String qName, Attributes attrs) throws SAXException {
            if (localName.equalsIgnoreCase(FORM_CONTROL_PR)) {
                String checkedValue = null;
                for (int i = 0; i < attrs.getLength(); i++) {
                    final String attrName = attrs.getLocalName(i);
                    if (attrName.equalsIgnoreCase(CHECKED)) {
                        checkedValue = attrs.getValue(i);
                        if (CHECKED.equalsIgnoreCase(checkedValue)) {
                            checked = true;
                        }
                        break;
                    }
                }
            }
        }

    }

    private static class DrawingCheckBoxTextContentHandler extends DefaultHandler {

        private static final String SP = "sp";
        private static final String CNVPR = "cNvPr";
        private static final String NAME = "name";
        private static final String TEXT = "t";

        private XlsxCheckBox checkBox;

        public DrawingCheckBoxTextContentHandler(XlsxCheckBox checkBox) {
            this.checkBox = checkBox;
        }

        private boolean inSp;
        private boolean inText;
        private String checkBoxText;

        /**
         * 解析 drawing#.xml 中的 checkbox 的信息
         * 其元素标签样例如下：
         *         <xdr:sp>
         *              ...
         *              <xdr:cNvPr hidden="1" name="Check Box 14" id="34830"> ... </xdr:cNvPr>
         *              ...
         *              <a:t>checkBox 文本内容</a:t>
         *              ...
         *         </xdr:sp>
         * @param uri
         * @param localName
         * @param qName
         * @param attrs
         * @throws SAXException
         */

        @Override
        public void startElement(String uri, String localName, String qName, Attributes attrs) throws SAXException {
            if (localName.equalsIgnoreCase(SP)) {
                inSp = true;
            }
            else if (inSp && localName.equalsIgnoreCase(CNVPR)) {

                for (int i = 0; i < attrs.getLength(); i++) {
                    final String attrName = attrs.getLocalName(i);
                    if (attrName.equalsIgnoreCase(NAME)) {   // 通过 checkBox 的名字确定标签获取文本内容
                        checkBoxText = attrs.getValue(i);
                        break;
                    }
                }
            }
            else if(inSp && localName.equalsIgnoreCase(TEXT)) {
                inText = true;
            }
        }

        /**
         * 解析标签的 text
         * @param ch
         * @param start
         * @param length
         * @throws SAXException
         */
        public void characters(char[] ch, int start, int length) throws SAXException {
            if (inText && Objects.equals(checkBoxText, checkBox.getName())) {
                // 当前标签是 <a:t>， 解析标签里的 text
                String value = new String(ch, start, length);
                checkBox.setText(value);
                checkBoxText = null;
            }
        }


        /**
         * 关闭标签，处理一些 flag
         * @param uri
         * @param localName
         * @param name
         * @throws SAXException
         */
        public void endElement(String uri, String localName, String name) throws SAXException {

            if (inSp && localName.equalsIgnoreCase(SP)) {
                inSp = false;
            }

            if (inText && localName.equalsIgnoreCase(TEXT)) {
                inText = false;
            }

        }

    }

    private static class ExtXSSReader extends XSSFReader {

        private OPCPackage pkg;
        private PackagePart workBookPart;

        public ExtXSSReader(OPCPackage pkg) throws IOException, OpenXML4JException {
            super(pkg);
            this.pkg = pkg;
            workBookPart  = pkg.getPartsByContentType(XSSFRelation.WORKBOOK.getContentType()).get(0);
        }

        public PackagePart getSheetPartById(String id) throws InvalidFormatException {
            URI refURI = workBookPart.getRelationship(id).getTargetURI();
            PackagePartName packagePartName = PackagingURIHelper.createPartName(refURI);
            return pkg.getPart(packagePartName);
        }

    }

    public static class XlsxCheckBox {

        private String name;
        private String rid;
        private int rowIndex;
        private int colIndex;
        private String text;

        private PackagePart ctrlPropPart;


        public String getName() {
            return name;
        }

        public void setName(String name) {
            this.name = name;
        }

        public String getRid() {
            return rid;
        }

        public void setRid(String rid) {
            this.rid = rid;
        }

        public int getRowIndex() {
            return rowIndex;
        }

        public void setRowIndex(int rowIndex) {
            this.rowIndex = rowIndex;
        }

        public int getColIndex() {
            return colIndex;
        }

        public void setColIndex(int colIndex) {
            this.colIndex = colIndex;
        }

        public String getText() {
            return text;
        }

        public void setText(String text) {
            this.text = text;
        }

        public PackagePart getCtrlPropPart() {
            return ctrlPropPart;
        }

        public void setCtrlPropPart(PackagePart ctrlPropPart) {
            this.ctrlPropPart = ctrlPropPart;
        }

        /**
         * 读取 ctrlProp 中的 checkBox 内容，获取 checked value
         * @return
         * @throws ParserConfigurationException
         * @throws SAXException
         * @throws IOException
         */
        public boolean getCheckedValue() throws ParserConfigurationException, SAXException, IOException {
            XMLReader reader  = XMLHelper.newXMLReader();

            // 构造 ctrlProp 处理器来解析标签
            // 由 reader 启动解析
            CtrlPropCheckBoxHandler handler = new CtrlPropCheckBoxHandler(this);
            reader.setContentHandler(handler);
            reader.parse(new InputSource(this.ctrlPropPart.getInputStream()));
            return handler.getChecked();
        }

        @Override
        public String toString() {
            return "XlsxCheckBox{" +
                    "rowIndex=" + rowIndex +
                    ", colIndex=" + colIndex +
                    ", text='" + text + '\'' +
                    '}';
        }
    }


    public static void main(String[] args) throws OpenXML4JException, ParserConfigurationException, IOException, SAXException {

        XlsxCheckBoxValueParser parser = new XlsxCheckBoxValueParser("info.xlsx");
        parser.chooseSheet("11-监管动态");
        System.out.println(parser.getCheckBoxValue(5, 3, "否"));

        parser.chooseSheet("1-企业基本情况");
        System.out.println(parser.getCheckBoxRows());
        System.out.println(parser.getCheckBoxCols(10));
        System.out.println(parser.getCheckBox(10, 2));
        System.out.println(parser.getCheckBox(10, 4));
        System.out.println(parser.getCheckBox(10, 5));

        System.out.println(parser.getCheckBoxValue(10, 5, "水处理"));

    }
}
