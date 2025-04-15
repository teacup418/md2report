# md4report

受md2report（下文缩写m2r）启发，重新构建了一个更灵活、更简单的项目。

相比m2r，m4r去除了大部分复杂的功能与配置，易于维护，上手难度低。

该版本仅封装为命令行应用，不会构建完整后端。

## 待办

- [ ] 构建命令行
- [ ] 构建配置保存与读取

## 项目结构说明


common：保存可以重复使用的通用文件

schools：保存各个学校个性化的文件

x/assets：logo等

x/config: 配置文件，例如实验报告、课程设计报告等

x/filters: pandoc过滤器，用于语义标签转化。

x/handlers: Python-docx代码，处理pandoc无法实现的功能。

x/ref: pandoc引用文档。

tests：测试用例

utils：保存可以重复使用的代码


config保存配置

- 名称
- 说明
- pandoc参数（如果有）
- handler流程（如果有）
- 作者信息、语录
- 赞助者信息、语录


templates保存模板

filter过滤器代码

handlers处理器代码



tests测试用例

## 运行流程说明

命令行调用m4r时，从命令行获取输入的参数。

读取目标md文件的元数据，并另外保存。

交给pandoc转换。

最后交由handler根据另存的元数据修改pandoc导出的文件。

## utils包功能说明

为段落添加边框。

在开头插入内容

在结尾插入内容

设置字体

设置颜色

## 配置开发环境

### 后端（Python）

配置一个虚拟环境，之后将在该环境中安装依赖，与其他项目隔离，方便单独维护。

```bash
python3 -m venv .venv
```
激活虚拟环境，Windows中运行

```powershell
.\.venv\Scripts\activate
```

然后安装依赖文件中的依赖

```bash
pip install -r requirements.txt
```

如果添加了新的依赖，将其安装的第三方库导出至`requirements.txt`

```bash
pip freeze > requirements.txt
```


```python
pip install python-frontmatter
```