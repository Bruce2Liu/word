package com.example.word.paper.bo;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

/**
 * 小题信息业务类
 * @author liujunhui
 * @date 2021/2/20 17:38
 */
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
public class SubQuestBO {
    /** 试题唯一标识 */
    private String id;

    /** 题干 */
    private String content;

    /** 试题序号，按整张试卷下从1开始排序 */
    private String questSeq;

    /** 试题序号，在完整题目下从1开始排序 */
    private String subQuestSeq;

    /** 试题题型 */
    private String questType;

    /** 试题分数 */
    private String questScore;

    /** 题目材料，eg：语文现代文阅读中的阅读材料 */
    private String materialOfContent;

    /** 题目做题要求 */
    private String expextationOfContent;

}
