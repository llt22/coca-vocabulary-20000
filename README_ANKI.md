# 如何将 COCA 词汇导入 Anki

本指南将帮助您将 COCA 20000 词汇导入 Anki 进行学习。

## 准备工作

1. 确保您已安装 [Node.js](https://nodejs.org/)（用于运行导出脚本）
2. 确保您已安装 [Anki](https://apps.ankiweb.net/)（用于记忆词汇）

## 导出步骤

1. 克隆或下载此仓库到本地
2. 打开命令行终端，进入仓库目录
3. 运行以下命令安装依赖：
   ```
   npm install
   ```
4. 运行导出脚本：
   ```
   node export_to_anki.js
   ```
5. 脚本将在仓库目录下创建一个 `anki_export` 文件夹，其中包含：
   - `all_words.csv`：包含所有 20000 个单词的文件
   - 多个分段的 CSV 文件（每个文件包含 1000 个单词）

## 导入 Anki

1. 打开 Anki 软件
2. 点击主界面的"导入文件"按钮
3. 选择 `anki_export` 文件夹中的任意 CSV 文件
   - 如果您想一次性导入所有单词，选择 `all_words.csv`
   - 如果您想分批导入，可以选择特定范围的文件，如 `part00_1-1000.csv`
4. 在导入设置中：
   - 确保字段映射正确（word, phonetic, definition, example, tags）
   - 选择合适的牌组（或创建新牌组）
   - 设置适当的导入选项
5. 点击"导入"按钮

## 自定义卡片模板

导入后，您可能需要自定义卡片模板以获得更好的学习体验：

1. 在 Anki 主界面点击"浏览"
2. 选择您导入的牌组
3. 点击"卡片"按钮
4. 自定义前模板和后模板，例如：

前模板示例：
```
<div class="word">{{word}}</div>
```

后模板示例：
```
<div class="word">{{word}}</div>
<div class="phonetic">{{phonetic}}</div>
<hr>
<div class="definition">{{definition}}</div>
<div class="example">{{example}}</div>
```

## 其他推荐的词汇学习软件

除了 Anki 外，还有其他一些优秀的词汇学习软件：

1. **Quizlet**：提供简单易用的界面和多种学习模式
2. **Memrise**：使用间隔重复系统和助记技巧帮助记忆
3. **SuperMemo**：使用科学的间隔重复算法
4. **百词斩**：专为中国用户设计的英语词汇学习软件
5. **扇贝单词**：提供丰富的例句和记忆方法

## 常见问题

**Q: 为什么导入后有些单词的音标显示不正确？**  
A: 可能是由于编码问题。尝试在 Anki 的导入设置中更改字符编码，或在卡片模板中使用适当的 CSS 样式。

**Q: 如何在 Anki 中添加单词发音？**  
A: 您可以使用 Anki 的插件，如 "AwesomeTTS" 来添加自动发音功能。

**Q: 如何根据词频分组学习？**  
A: CSV 文件中的 tags 字段已包含词频范围信息，您可以在 Anki 中使用标签进行筛选和学习。