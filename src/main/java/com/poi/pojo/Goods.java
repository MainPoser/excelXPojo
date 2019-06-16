package com.poi.pojo;

import java.io.Serializable;
import java.math.BigDecimal;

/**
 * @author hallele
 */

public class Goods implements Serializable {
    /**
     * 主键
     */
    private Long id;

    /**
     * 商家ID
     */

    private String sellerId;

    /**
     * SPU名
     */

    private String goodsName;

    /**
     * 默认SKU
     */

    private Long defaultItemId;

    /**
     * 状态 表示商品是否被审计 0 未审核 1 已审核 2 审核未通过 3 关闭
     */

    private String auditStatus;

    /**
     * 是否上架
     */

    private String isMarketable;

    /**
     * 品牌
     */

    private Long brandId;

    /**
     * 副标题
     */
    private String caption;

    /**
     * 一级类目
     */

    private Long category1Id;

    /**
     * 二级类目
     */

    private Long category2Id;

    /**
     * 三级类目
     */

    private Long category3Id;

    /**
     * 小图
     */

    private String smallPic;

    /**
     * 商城价
     */
    private BigDecimal price;

    /**
     * 分类模板ID
     */

    private Long typeTemplateId;

    /**
     * 是否启用规格
     */

    private String isEnableSpec;

    /**
     * 是否删除
     */

    private Boolean isDelete;

    private static final long serialVersionUID = 1L;

    /**
     * 获取主键
     *
     * @return id - 主键
     */
    public Long getId() {
        return id;
    }

    /**
     * 设置主键
     *
     * @param id 主键
     */
    public void setId(Long id) {
        this.id = id;
    }

    /**
     * 获取商家ID
     *
     * @return seller_id - 商家ID
     */
    public String getSellerId() {
        return sellerId;
    }

    /**
     * 设置商家ID
     *
     * @param sellerId 商家ID
     */
    public void setSellerId(String sellerId) {
        this.sellerId = sellerId;
    }

    /**
     * 获取SPU名
     *
     * @return goods_name - SPU名
     */
    public String getGoodsName() {
        return goodsName;
    }

    /**
     * 设置SPU名
     *
     * @param goodsName SPU名
     */
    public void setGoodsName(String goodsName) {
        this.goodsName = goodsName;
    }

    /**
     * 获取默认SKU
     *
     * @return default_item_id - 默认SKU
     */
    public Long getDefaultItemId() {
        return defaultItemId;
    }

    /**
     * 设置默认SKU
     *
     * @param defaultItemId 默认SKU
     */
    public void setDefaultItemId(Long defaultItemId) {
        this.defaultItemId = defaultItemId;
    }

    /**
     * 获取状态 表示商品是否被审计 0 未审核 1 已审核 2 审核未通过 3 关闭
     *
     * @return audit_status - 状态 表示商品是否被审计 0 未审核 1 已审核 2 审核未通过 3 关闭
     */
    public String getAuditStatus() {
        return auditStatus;
    }

    /**
     * 设置状态 表示商品是否被审计 0 未审核 1 已审核 2 审核未通过 3 关闭
     *
     * @param auditStatus 状态 表示商品是否被审计 0 未审核 1 已审核 2 审核未通过 3 关闭
     */
    public void setAuditStatus(String auditStatus) {
        this.auditStatus = auditStatus;
    }

    /**
     * 获取是否上架
     *
     * @return is_marketable - 是否上架
     */
    public String getIsMarketable() {
        return isMarketable;
    }

    /**
     * 设置是否上架
     *
     * @param isMarketable 是否上架
     */
    public void setIsMarketable(String isMarketable) {
        this.isMarketable = isMarketable;
    }

    /**
     * 获取品牌
     *
     * @return brand_id - 品牌
     */
    public Long getBrandId() {
        return brandId;
    }

    /**
     * 设置品牌
     *
     * @param brandId 品牌
     */
    public void setBrandId(Long brandId) {
        this.brandId = brandId;
    }

    /**
     * 获取副标题
     *
     * @return caption - 副标题
     */
    public String getCaption() {
        return caption;
    }

    /**
     * 设置副标题
     *
     * @param caption 副标题
     */
    public void setCaption(String caption) {
        this.caption = caption;
    }

    /**
     * 获取一级类目
     *
     * @return category1_id - 一级类目
     */
    public Long getCategory1Id() {
        return category1Id;
    }

    /**
     * 设置一级类目
     *
     * @param category1Id 一级类目
     */
    public void setCategory1Id(Long category1Id) {
        this.category1Id = category1Id;
    }

    /**
     * 获取二级类目
     *
     * @return category2_id - 二级类目
     */
    public Long getCategory2Id() {
        return category2Id;
    }

    /**
     * 设置二级类目
     *
     * @param category2Id 二级类目
     */
    public void setCategory2Id(Long category2Id) {
        this.category2Id = category2Id;
    }

    /**
     * 获取三级类目
     *
     * @return category3_id - 三级类目
     */
    public Long getCategory3Id() {
        return category3Id;
    }

    /**
     * 设置三级类目
     *
     * @param category3Id 三级类目
     */
    public void setCategory3Id(Long category3Id) {
        this.category3Id = category3Id;
    }

    /**
     * 获取小图
     *
     * @return small_pic - 小图
     */
    public String getSmallPic() {
        return smallPic;
    }

    /**
     * 设置小图
     *
     * @param smallPic 小图
     */
    public void setSmallPic(String smallPic) {
        this.smallPic = smallPic;
    }

    /**
     * 获取商城价
     *
     * @return price - 商城价
     */
    public BigDecimal getPrice() {
        return price;
    }

    /**
     * 设置商城价
     *
     * @param price 商城价
     */
    public void setPrice(BigDecimal price) {
        this.price = price;
    }

    /**
     * 获取分类模板ID
     *
     * @return type_template_id - 分类模板ID
     */
    public Long getTypeTemplateId() {
        return typeTemplateId;
    }

    /**
     * 设置分类模板ID
     *
     * @param typeTemplateId 分类模板ID
     */
    public void setTypeTemplateId(Long typeTemplateId) {
        this.typeTemplateId = typeTemplateId;
    }

    /**
     * 获取是否启用规格
     *
     * @return is_enable_spec - 是否启用规格
     */
    public String getIsEnableSpec() {
        return isEnableSpec;
    }

    /**
     * 设置是否启用规格
     *
     * @param isEnableSpec 是否启用规格
     */
    public void setIsEnableSpec(String isEnableSpec) {
        this.isEnableSpec = isEnableSpec;
    }

    /**
     * 获取是否删除
     *
     * @return is_delete - 是否删除
     */
    public Boolean getIsDelete() {
        return isDelete;
    }

    /**
     * 设置是否删除
     *
     * @param isDelete 是否删除
     */
    public void setIsDelete(Boolean isDelete) {
        this.isDelete = isDelete;
    }

    @Override
    public String toString() {
        StringBuilder sb = new StringBuilder();
        sb.append(getClass().getSimpleName());
        sb.append(" [");
        sb.append("Hash = ").append(hashCode());
        sb.append(", id=").append(id);
        sb.append(", sellerId=").append(sellerId);
        sb.append(", goodsName=").append(goodsName);
        sb.append(", defaultItemId=").append(defaultItemId);
        sb.append(", auditStatus=").append(auditStatus);
        sb.append(", isMarketable=").append(isMarketable);
        sb.append(", brandId=").append(brandId);
        sb.append(", caption=").append(caption);
        sb.append(", category1Id=").append(category1Id);
        sb.append(", category2Id=").append(category2Id);
        sb.append(", category3Id=").append(category3Id);
        sb.append(", smallPic=").append(smallPic);
        sb.append(", price=").append(price);
        sb.append(", typeTemplateId=").append(typeTemplateId);
        sb.append(", isEnableSpec=").append(isEnableSpec);
        sb.append(", isDelete=").append(isDelete);
        sb.append(", serialVersionUID=").append(serialVersionUID);
        sb.append("]");
        return sb.toString();
    }
}