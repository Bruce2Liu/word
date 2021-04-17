## 技术

- aspose word for java

    - 用于获取相应的ole对象以及转换后latex公式写回

    - 用于获取OMML公式以及转换后latex公式写回

- ruby
    - 用于ole文件转换xml

- xslt/xproc
    - 用于xml文件转换mathnl

- fmath
    - 用于mathml转换latex


`注意的地方：`

1. 在执行ruby原生代码的时候，1000个公式进行转换使用10秒钟，而在使用java调用ruby代码的时候1000个公式使用30秒钟。（执行的ruby测试代码完全相同）

2. 在使用xslt转换xml1的时候，是通过遍历xml list一个一个转换的，而不是一次加载完所有的xml文件，可优化。（目前1000个公式需要3分钟，测试了半小时，未见出现OOM问题）

3. 在图片中存在公式的现象，即公式和图片进行组合，这样就会导致在获取相应公式后转换完成无法写回，因为文字和图片无法组合。<br>
[四川省自贡市富顺县赵化中学2017-2018学年度下期八年级数学_数据的分析_单元训练题__无答案_.docx](/uploads/d1bfacea0c2e7a7ed45c78bc1190b6b0/四川省自贡市富顺县赵化中学2017-2018学年度下期八年级数学_数据的分析_单元训练题__无答案_.docx)

4. 在获取公式的时候，会判断progId.contains("Equation")，但是存在一种情况，如果progid为空，有时候也是公式，只是在word中是图片的形式，但依旧是可以转化的，目前已兼容，但是否存在问题，需要进一步测试。<br/>
[2015浙江衢州中考数学卷.docx](/uploads/996b2f0c0d558626246184e78b924aa4/2015浙江衢州中考数学卷.docx)


## 接口说明
输入一个word文件，如果里面含有公式，会输出一个word文件，其中里面的公式被转换为了latex格式。<br/>
之所以要转化为latex格式,是因为后续分析需要使用。

## 接口名
word2Latex


## 请求参数说明

| 参数 | 说明 | 参数类型|是否必填 | 校验规则 |
| -------- | -------- | -------- | -------- | -------- |
| wordPath | 原始文件绝对路径  | String | 是  |  |
| targetPath | 生成文件的绝对路径  | String | 否  |  |


## 返回参数说明

false：公式转换失败。

    原因：
    1. mathtype公式在转化为xml的时候，转换失败。
      - 公式版本不支持
      - 公式在本地打不开
    2. xml转换mathml的时候失败。
    3. 写回latex公式的时候失败。
    4. omml公式转换mathml时失败。


true：转换成功


# 20200425公式实现方式流程

目的：去除aspose word使用，优化转换正确率

- mathtype：
    - mathtype到xml再到mathml流程没有改
    - mathml到latex改为使用mmltex.xsl方式实现
    - 解析mathtype从aspose word使用docx4j代替

- omml：
    - 从omml到mathml，使用omml2mml.xsl方式替代
    - mathml到latex改为使用mmltex.xsl方式实现
    - 解析omml从aspose word使用docx4j代替


```
判断是否为公式的条件：
通过读取rels文件中的信息，获取类型为oleObject类型的关联信息，之后获取对应的文件名以及rId。
怎样获取公式：
通过rId获取对应的Part，之后转为字节流的Part，之后获取字节流
怎样定位document.xml中公式的位置：
目前公式是在r标签中。
但是可能存在表格，等标签中，需要考虑。
本来考虑一级一级遍历，后来改为直接递归获取所有的公式Object，之后转换代替。
怎样替换：
需要找到对应公式所在的R或者P标签，之后进行set替换。

域公式的处理也放到解析流程里来

```

# 目前公式存在地方：

- mathtype：
R标签中

- omml：
P标签中，和R同级（不论是oMathPara还是oMath）

而R标签一定存在于P标签中。

P标签有可能存在表格，图片中。



[图片中有公式.docx](/uploads/f9a1f46313b39c6b3c1ecbc43cb352dc/图片中有公式.docx)

[0fbb6752ec3c454681c66a1a4cd74f2b.docx](/uploads/cf467c48c8573524e176468c9cba7985/0fbb6752ec3c454681c66a1a4cd74f2b.docx)







