package com.example.word.paper.bo;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.util.List;

/**
 * 试题信息业务类
 * @author liujunhui
 * @date 2021/2/20 17:37
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
public class QuestBO {
    /** 试题唯一标识 */
    private String id;

    /** 题干 */
    private String content;

    /** 大题描述 */
    private String stage;

    /** 二级标题，eg：语文课本单元题的“字词句运用” */
    private String partName;

    /** 试题题型 */
    private String questType;

    /** 试题序号 */
    private String questSeq;

    /** 试题难度 */
    private String difficulty;

    /** 试题分数 */
    private String questScore;

    /** 试题答案 */
    private String answer;

    /** 试题答案解析 */
    private String analysis;

    /** 知识点 */
    private String points;

    /** 题目材料，eg：语文现代文阅读中的阅读材料 */
    private String materialOfContent;

    /** 题目做题要求 */
    private String expextationOfContent;

    /** 领域 */
    private String field;

    /** 题目在试卷中的位置 */
    private String questLocation;

    /** 文件名 */
    private String docname;

    private String resourcename;

    private String resourceType;

    /** 试卷Id */
    private String paperId;

    /** 生成小word试题路径 todo 确认不再用的话可以去掉 */
    private String qPath;

    /** 试题选项 */
    private String choices;

    /** 小题 */
    private List<SubQuestBO> subQuestBOList;
}
