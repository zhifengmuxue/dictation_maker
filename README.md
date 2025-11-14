# Word 红字转默写版脚本

这是一个小脚本，用于把 Word 文档（.docx）中带有**显式颜色（有 RGB 值）且不是黑色**的文字替换为等长的下划线占位（通过下划线样式的不可断行空格实现），便于制作默写试题。

依赖
- Python >= 3.7
- python-docx

安装依赖
```bash
uv install
```

用法（命令行）
```bash
# 在项目目录或任意目录下运行
python make_dictation.py path/to/input.docx
```
示例（macOS）
```bash
cd /Users/lima/code/language/python/script/dictation_maker
python make_dictation.py ~/Documents/example.docx
# 输出文件会生成在同目录： ~/Documents/example_dictation.docx
```

行为说明
- 脚本通过检查每个 run 的 run.font.color.rgb 来判断是否替换：
  - 当 run.font.color 存在且 rgb 不为 None 且不等于 RGB(0,0,0) 时，视为“非黑色”，该 run 的文本会被替换为等长的不可断行空格（\u00A0），并设置 run.font.underline = True，颜色改为黑色。
  - 如果 run 没有显式颜色（color 为 None）或没有 RGB 信息（例如 theme color），脚本保守地不替换该文本。
- 输出文件命名规则：如果输入为 `input.docx`，则输出 `input_dictation.docx`（与输入同目录）。

脚本限制
- 当前只检测 run 级别的显式 RGB 颜色（非黑即替换），不进行颜色近似匹配。
- 未处理页眉/页脚、文本框/形状中的文字或基于 theme 的颜色。
- 尽量保留原有样式，但可能会修改字体颜色与下划线属性。

改进方向
- 增加对红色近似的阈值匹配（例如只替换接近红色的颜色）。
- 支持页眉/页脚、文本框与形状内文本、以及 theme 色的解析与处理。
- 增加命令行选项：指定输出路径、选择颜色匹配策略（严格 RGB / 阈值 / 指定颜色）等。
