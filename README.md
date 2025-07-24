# CVE数据抓取与百炼大模型分析工具

本工具可自动抓取指定日期范围内的CVE高危漏洞，生成格式化Excel报告，并一键上传到百炼平台，调用大模型自动分析，适合安全研究、漏洞管理等场景。

---

## 主要功能

- **交互式日期选择**，批量抓取多个组件的CVE高危漏洞（CVSS≥7.0）。
- **多线程并发抓取**，大幅提升数据获取速度。
- **自动生成美观Excel报告**，组件名称自动合并单元格，含“漏洞级别”列。
- **一键上传Excel到百炼平台**，自动读取`prompt.txt`作为分析提示词，调用大模型分析。
- **分析结果自动保存为 `bailian_analysis.txt`**。
- **分析后自动清理百炼平台所有已上传文件，保障云端空间整洁**。

---

## 快速开始

### 1. 环境准备

- Python 3.8+
- 推荐使用虚拟环境

### 2. 安装依赖

```bash
pip install requests pandas openpyxl openai
```

### 3. 配置文件说明

- `components.txt`  
  每行一个组件名，示例：
  ```
  nginx
  mysql
  elasticsearch
  spring
  ```
- `key.txt`  
  填写你的百炼平台API Key，**仅一行**，无空格、无BOM。
- `prompt.txt`  
  填写你的大模型分析提示词（如“请帮我检查Excel中是否有不匹配的CVE编号”）。

### 4. 使用方法

1. **运行主程序：**
   ```bash
   python cve_workflow.py
   ```
2. **按提示输入日期范围，确认抓取。**
3. **抓取完成后，自动生成Excel报告。**
4. **程序会询问是否上传并用大模型分析，输入`y`即可自动上传、分析并保存结果。**

---

## 输出文件

- `cve_results_YYYYMMDD_HHMMSS.xlsx`  
  自动生成的漏洞报告，含组件名称、CVE编号、发布时间、漏洞描述、漏洞级别。
- `bailian_analysis.txt`  
  百炼大模型分析结果。

---

## 注意事项

- `key.txt`、`prompt.txt`、`components.txt`请勿上传到公共仓库。
- 百炼平台API Key请妥善保管。
- 若需自定义并发线程数、Excel样式等，可直接修改`cve_workflow.py`相关参数。

---

## 常见问题

- **编码报错**：请确保`key.txt`、`prompt.txt`均为UTF-8编码，无BOM。
- **openai库未安装**：请先`pip install openai`。
- **API Key无效**：请联系百炼平台获取有效Key。

---

## 许可证

MIT License

---

如有问题或建议，欢迎反馈！ 