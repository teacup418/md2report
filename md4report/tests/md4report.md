# 需求
易于更新，通过配置表可以高度复用组件。

# 依赖
pyyaml
`pip install PyYAML`

命令
```powershell
pandoc -s "C:\Users\Vincent\Documents\code\md2report\md4report\tests\teacup.docx" --reference-doc "C:\Users\Vincent\Documents\code\md2report\md4report\templates\GZMTU.docx" --filter "C:\Users\Vincent\Documents\code\md2report\md4report\src\core\filters\general.py" -o "C:\Users\Vincent\Documents\code\md2report\md4report\tests\teacup.docx"
```

# 流程
1. 使用pandoc生成docx文件
2. 在docx文件添加标题、成绩、学生信息


# 添加信息
到首个段落之前
添加标题
标题后添加成绩
成绩后添加学生信息


# 流程
引用文档
filter



# 项目架构
```kotlin
fun largerNumber(num1: Int, num2: Int): Int {
	val value = if (num1 > num2) {
		num1
	} else {
		num2
	}
	return value
}
```

# 参考文献
[lxml](https://lxml.de/tutorial.html)
[贡献手册](https://woolen-sheep.github.io/md2report/contribute/)
[代码仓库](https://github.com/woolen-sheep/md2report)
[pandoc 过滤器](https://pandoc.org/filters.html)
[[panflute]] pandoc的Python封装
[pandoc参数说明](https://pandoc.org/MANUAL.html)尤其是`--reference-doc`
[[python-docx]]
[微软XML格式API文档](https://learn.microsoft.com/en-us/dotnet/api/overview/openxml/?view=openxml-3.0.1)
[文档预览JS库](https://gitee.com/kekingcn/file-online-preview/wikis/pages)
[中间插入](https://github.com/python-openxml/python-docx/issues/156)