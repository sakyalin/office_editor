# Office文档错别字检查工具

这是一个使用OpenAI API检查Office文档（Word、PowerPoint、Excel）中错别字和病句的工具。它会在原文后面添加红色文字显示的修改建议，并生成一个带"修订"后缀的新文件。

## 功能特点

- 支持检查Word文档（.docx）中的错别字和病句
- 支持检查PowerPoint演示文稿（.pptx）中的错别字和病句
- 支持检查Excel工作簿（.xlsx）中的错别字和病句
- 使用OpenAI API进行智能文本检查
- 在原文后面添加红色文字显示的修改建议
- 生成带"修订"后缀的新文件，不修改原始文件
- 简洁直观的用户界面

## 安装依赖

```bash
pip install -r requirements.txt
```

## 使用方法

1. 配置OpenAI API密钥：

   在项目根目录下的`.env`文件中设置您的API密钥：

   ```
   OPENAI_API_KEY=your_api_key_here
   ```

   将`your_api_key_here`替换为您的实际OpenAI API密钥。

2. 运行程序：

   ```bash
   python main.py
   ```

3. 在界面中点击"浏览"按钮，选择要检查的Office文档。
4. 点击"开始处理"按钮，等待处理完成。
5. 处理完成后，将在原文件所在目录生成一个带"修订"后缀的新文件。

## 打包为可执行文件

### Windows

```bash
./convert_icon.sh
pyinstaller --name="Office文档错别字检查工具" --windowed --icon=icon.ico --add-data="icon.ico;." main.py
```

### macOS

```bash
./convert_icon.sh
pyinstaller --name="Office文档错别字检查工具" --windowed --icon=icon.icns --add-data="icon.icns:." main.py
```

## 注意事项

- 需要在`.env`文件中配置有效的OpenAI API密钥才能使用此工具。
- 如果`.env`文件不存在或API密钥未设置，程序将无法正常工作。
- 处理大型文件可能需要较长时间，请耐心等待。
- 使用的模型为 `gpt-4o-mini`，确保您的API密钥有权限访问此模型。
- 如果API调用失败，程序会使用简单的模拟数据作为备选，以便您可以测试其他功能。

## 许可证

MIT