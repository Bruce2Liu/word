package com.example.word.paper.bo;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.util.List;
import java.util.Map;

/**
 * 试卷信息业务类
 * @author liujunhui
 * @date 2021/2/20 17:37
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
public class PaperBO {
    /** 试卷id */
    private String paperId;

    /** 资源名（考试名称）；试卷文件最上方考试名称信息 */
    private String resourceName;

    /** 学科编码 */
    private String subject;

    /** 版本编码 */
    private String version;

    /** 学段 */
    private String phase;

    /** 学年 */
    private String academicYear;

    /** 年份 */
    private String year;

    /** 年级 */
    private String grade;

    /** 学期 */
    private String term;

    /** 单元名称 */
    private String unitName;

    /** 单元序号 */
    private String unitSeq;

    /** 单元下小节名称 */
    private String section;

    /** 单元下小节序号 */
    private String sectionSeq;

    /** 试题数量 */
    private String questNum;

    /** 试卷满分 */
    private String fullMarks;

    /** 考试时长 */
    private String testTime;

    /** 省份 */
    private String pro;

    /** 地区 */
    private String area;

    /** 县区 */
    private String subArea;

    /** 学校 */
    private String orgId;

    /** 资源类型 */
    private String resourceType;

    /** 子类型 */
    private String subType;

    /** 试卷类型 */
    private String orgStage;

    /** 试卷状态 */
    private String status;

    /** 导入时间 */
    private String importDate;

    /** 试卷json格式数据在hdfs上的文件路径 */
    private String path;

    /** 采集来源 */
    private String collectionSource;

    /** 课本标识 */
    private String bookId;

    /** 试卷所有试题信息（试题序号-试题id-试题分数;） */
    private String quests;

    /** 原试卷在hdfs上的路径 */
    private String paperPath;

    /** 试卷导入路径 */
    private String importPath;

    /** 图片 */
    private Map<String, byte[]> imgMap;

    /** 试卷小word */
    //private Map<String, byte[]> smallWordMap;

    /** 试卷所有试题 */
    private List<QuestBO> htmlQuestBOList;
}
