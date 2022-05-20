package org.example.word;

import com.alibaba.fastjson.JSONObject;
import freemarker.template.Configuration;
import freemarker.template.Template;
import freemarker.template.Version;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @author xxd
 * @date 2022/5/20 16:25
 */
public class Test {

    public static void main(String[] args) throws Exception {
        List<JSONObject> headList = new ArrayList<>();
        JSONObject head = new JSONObject();
        head.put("name", "气量管理");
        head.put("value", "（15%）");
        headList.add(head);
        head = new JSONObject();
        head.put("name", "输差管理");
        head.put("value", "（5%）");
        headList.add(head);
        head = new JSONObject();
        head.put("name", "生产调度");
        head.put("value", "（15%）");
        headList.add(head);
        head = new JSONObject();
        head.put("name", "应急管理");
        head.put("value", "（15%）");
        headList.add(head);
        head = new JSONObject();
        head.put("name", "运行管理");
        head.put("value", "（25%）");
        headList.add(head);
        head = new JSONObject();
        head.put("name", "设备管理");
        head.put("value", "（20%）");
        headList.add(head);
        head = new JSONObject();
        head.put("name", "规划管理");
        head.put("value", "（5%）");
        headList.add(head);

        List<JSONObject> dataList = new ArrayList<>();
        JSONObject data = new JSONObject();
        data.put("num", 1);
        data.put("totalScore", 94);
        data.put("company", "慈溪公司");
        List<String> headDataList = new ArrayList<>();
        headDataList.add("99.00");
        headDataList.add("98.00");
        headDataList.add("92.00");
        headDataList.add("92.86");
        headDataList.add("94.21");
        headDataList.add("94.67");
        headDataList.add("100.00");
        data.put("headDataList", headDataList);
        dataList.add(data);
        data = new JSONObject();
        data.put("num", 1);
        data.put("totalScore", 94);
        data.put("company", "慈溪公司");
        headDataList = new ArrayList<>();
        headDataList.add("97.00");
        headDataList.add("96.00");
        headDataList.add("92.00");
        headDataList.add("92.86");
        headDataList.add("94.21");
        headDataList.add("94.67");
        headDataList.add("100.00");
        data.put("headDataList", headDataList);
        dataList.add(data);

        List<String> avgList = new ArrayList<>();
        avgList.add("99.00");
        avgList.add("98.00");
        avgList.add("92.00");
        avgList.add("92.86");
        avgList.add("94.21");
        avgList.add("94.67");
        avgList.add("100.00");

        Map<String, Object> params = new HashMap<>();
        params.put("headList", headList);
        params.put("dataList", dataList);
        params.put("avgTotalScore", "97.83");
        params.put("avgList", avgList);
        String fileName="test.doc";
        generateWord(params, "src/main/resources/temp/" + fileName, "src/main/resources/word/", "template.xml");
    }

    /**
     * 使用FreeMarker自动生成Word文档
     *
     * @param dataMap  生成Word文档所需要的数据
     * @param fileName 生成Word文档的全路径名称
     */
    public static void generateWord(Map<String, Object> dataMap, String fileName, String templatePath, String template) throws Exception {
        // 设置FreeMarker的版本和编码格式
        Configuration configuration = new Configuration(new Version("2.3.23"));
        configuration.setDefaultEncoding("UTF-8");

        // 设置FreeMarker生成Word文档所需要的模板的路径
        configuration.setDirectoryForTemplateLoading(new File(templatePath));
        // 设置FreeMarker生成Word文档所需要的模板
        Template t = configuration.getTemplate(template, "UTF-8");
        // 创建一个Word文档的输出流
        Writer out = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(new File(fileName)), StandardCharsets.UTF_8));
        //FreeMarker使用Word模板和数据生成Word文档
        t.process(dataMap, out);
        out.flush();
        out.close();
    }
}
