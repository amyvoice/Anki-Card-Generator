# Anki卡片生成器 - LED显示屏版
这是一款简单的 **Anki 卡片生成工具**，支持输入英文和中文内容，能**自动生成 PPT（PowerPoint）文件**，方便你批量制作 Anki 单词卡。

## 🔧 功能简介

- 用户可在软件界面自己手动输入英文和中文，一张张保存为卡片
- 高频操作：保存并继续写（SAVE）、生成 PPT（GEN PPT）、清空卡片（CLEAR）
- 软件界面的数字模拟像素显示屏风格
- 生成的 PPT 每张卡片包含：
  - 英文（Courier New 48pt 粗体）
  - 中文（微软雅黑 38pt 稍小字体）
- 如有批量制作PPT卡片的需求，可以从 **TXT 或 MD 文件** 中批量读取英语和中文词汇对，自动生成卡片并保存为同名的 PPT 和 TXT 文件

- TXT或MD文件的模板如下，可复制后用于制作批量卡片专用文本：

=== 第1页 ===
英文: 【可在这里填写英文句子】
中文: 【可在这里填写中文翻译】

=== 第2页 ===
英文: I'll get back to you on that.
中文: 我稍后回复你。

=== 第3页 ===
英文: Do you have a minute?
中文: 你有空吗？

---

## 📝 如何使用？

### 📝 在电脑上运行（Python 环境）

1. 确保你的电脑已安装 Python 和相关库：

   ```bash
   pip install python-pptx

2. 下载并运行 AnkiCard.py 文件，就看到一个窗口弹出，可以直接输入英文和中文词汇。（也可使用bat文件）
3. 点击 SAVE，保存当前页，并继续写下一页，文本框自动清空。
4. 如果卡片内容都输入好了，可点击 GEN PPT，会在与py文件相同路径下，自动生成一个【Anki年月日时分.pptx】以及一个同名txt 文件，方便备份卡片文本。
5. 点击 CLEAR，可清空所有卡片，重新开始写。

<img width="80%" height="80%" alt="image" src="https://github.com/user-attachments/assets/ab79234f-cc5f-443c-848c-b7be5c27adb5" />

<img width="80%" height="80%" alt="image" src="https://github.com/user-attachments/assets/d356c7e8-9bda-4252-a92a-f8945a7c9b5f" />

<img width="80%" height="80%" alt="image" src="https://github.com/user-attachments/assets/cee98c9b-dddb-4830-9833-cbae84549601" />

