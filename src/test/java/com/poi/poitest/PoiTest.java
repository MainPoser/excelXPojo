package com.poi.poitest;


import com.poi.pojo.Brand;
import com.poi.pojo.Goods;
import com.poi.util.ExcelUtil;
import org.junit.Test;

import java.io.*;
import java.net.URL;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;

/**
 * @author:tianyao
 * @date:2019-06-15 16:57
 */
public class PoiTest {
    /**
     * pojo属性转换成excel表格测试
     * @throws Exception
     */
    @Test
    public  void pojo2Excel() throws Exception {
        //获得list集合
        List<Brand> list = new ArrayList<>();
        Brand brand = new Brand();
        //id默认不显示，因为正常数据库的数据的 id不会打印出来，这里只是测试
        brand.setId(1);
        brand.setName("测试");
        brand.setFirstChar("C");
        list.add(brand);
        //新建一个输出流
        File file = new File("src/main/resources/brand.xlsx");
        FileOutputStream outputStream = new FileOutputStream(file);
        //创建一个表名和pojo属性名的 对应关系map
        LinkedHashMap<String, String> map = new LinkedHashMap<>();
        map.put("name","品牌名");
        map.put("firstChar","首字母大写");
        //创建表头
        String headLine =  "品牌表";
        //属性列数-1(列是从0开始的)
        int size = map.size()-1;
        ExcelUtil.pojo2Excel(list,outputStream,map,headLine,size);
    }

    /**
     * 表格转换成pojo类测试
     */
    @Test
    public void excel2Pojo() throws IOException {
        //创建一个输入流
        File file = new File("src/main/resources/goods.xlsx");
        FileInputStream fileInputStream = new FileInputStream(file);
        //获取到要转换的pojo的class对象
        Class<Goods> clazz = Goods.class;
        //表格中的列名和pojo属性名的对应关系
        LinkedHashMap<String, String> map = new LinkedHashMap<>();
        map.put("卖家名","sellerId");
        map.put("商品名","goodsName");
        map.put("默认分类名","defaultItemId");
        map.put("状态","auditStatus");
        map.put("是否mark","isMarketable");
        map.put("品牌id","brandId");
        map.put("副标题","caption");
        map.put("一级菜单","category1Id");
        map.put("二级菜单","category2Id");
        map.put("三级菜单","category3Id");
        map.put("图片链接","smallPic");
        map.put("价格","price");
        map.put("模板id","typeTemplateId");
        map.put("是否启用spec","isEnableSpec");
        map.put("是否删除","isDelete");
        List<Goods> goods = ExcelUtil.excel2Pojo(fileInputStream, clazz, map);
        System.out.println(goods);
    }
}
