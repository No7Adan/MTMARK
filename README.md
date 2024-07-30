# MTMARK
自研翻译质量评估工具MTMARK不仅支持结合参考译文的多段文本并行分析，同时在自动输出每段BLEU值和MQM评价报告的同时，还可以完成输出数据的可视化，使得翻译质量评估的结果更加直观和全面，更方便使用者平行对比多个机器翻译质量。
此次实践也刻意使用语言结构较为复杂、词汇难度较高的经济学人相关语料，较为科学且全面地展示了该自研工具的可用性和可靠性。同时在实践过程中我们也发现，除去MQM分析结果由于GPT生成内容的随机性偶尔会出现错误的情况出现，使用该工具对机器翻译质量评估，最终所得到的判断结果同人工译者审校后所得到的结果的相似度十分之高。

## 作者
北京语言大学 翻译（本地化）专业 2021级 綦昊、项一卓、郑楚仪、张思怡

## 对所使用翻译评价标准（GEMBA-MQM, BLEU）的说明
（1）GEMBA-MQM
在使用MQM评估机器翻译质量时，出现了一些明显的评估错误和不准确之处。此次实践主要基于GPT-4o进行评估，而因模型其本身具有一定的不确定性，我们发现我们最后所得到的评估结果也存在一定的不确定性，虽然这种不确定性导致的错误产生的概率极小，且时常不受我们控制，但我们也应该意识到问题的存在。在我们的实践中，所遇到的主要错误类型包括但不限于以下几种：
- 错译评估错误：MQM经常凭空认为译文存在错译，特别是在经济、金融等特定领域术语和一些复杂表达的评估上。系统常常将正确翻译误认为错误，从而导致MQM分数奇低的现象出现。
- 漏译评估错误：MQM在多次评估中错误地认为译文漏译了某些内容，尽管实际译文中已经正确翻译了这些部分。
- 语义理解错误：系统对一些复杂语句的语义理解出现偏差，导致将一些细小的错误误判为严重错误的情况出现。
由于GPT模型的生成不确定性，MQM在评估结果上表现出一定的不一致性。同样的句子在不同评估中可能得出不同结论。总体而言，MQM在处理含有专业术语和复杂句法的语句时表现较差，评估结果需要人工去进行进一步的优化，以提高其可靠性和准确性。

（2）BLEU
BLEU评估方法是一种常用的机器翻译质量评估指标，主要衡量译文与参考译文之间的相似度。BLEU通过计算译文和参考译文中n-gram（n元语法）匹配的数量来评估翻译质量，其中n-gram可以是单词、词组或短语。
研究者们已经发现，BLEU值与译文的流畅性、词汇和短语的精确匹配度相关度较大。这意味着，译文中使用的词汇和短语与参考译文越相似，BLEU值就越高。然而，BLEU并不直接评估译文的语义准确性和连贯性，只是通过表面词汇的匹配来间接反映翻译质量。
结合我们的此次实践的结果不难发现，BLEU值在衡量译文流畅性方面表现出较高的相关性，这与我们使用MTMARK工具得出的结论一致。因此，BLEU值可以作为评估译文质量的一个有效指标，尤其是在评估流畅性和匹配度方面。虽然BLEU在某些情况下可能无法全面反映译文的质量，但它依然具有一定的参考价值，依然可以结合其它质量评估标准对译文质量进行多维度的评估，从而帮助研究者们更准确地衡量机器翻译工具的性能。

## 使用说明
1.	参考[fix: use Fraction override · nltk/nltk@86fa083](https://github.com/nltk/nltk/commit/86fa0832f0f4b366f96867f59ae05d744d68b513)对NLTK库原生BLEU进行修复。
2.	确保utils解压在main.py的相同目录下，以运行有道翻译API。
3.	如果输出报告不在当前程序目录下，请前往用户文件夹。
4.	因为Python Docx库相关问题，可能遇到显示字体错误，请更换Word字体。
5.	请将所有的以APIKEY开头的字符串更换为自己的API密钥。
