#!/bin/bash

# 检查是否安装了必要的工具
if ! command -v sips &> /dev/null; then
    echo "错误: 需要sips工具 (应该已经在macOS中预装)"
    exit 1
fi

if ! command -v iconutil &> /dev/null; then
    echo "错误: 需要iconutil工具 (应该已经在macOS中预装)"
    exit 1
fi

# 首先将SVG转换为高分辨率PNG
sips -s format png icon.svg --out icon_large.png

# 创建临时iconset目录
ICONSET="icon.iconset"
mkdir -p "$ICONSET"

# 使用sips生成不同尺寸的PNG
sips -z 16 16 icon_large.png --out "$ICONSET/icon_16x16.png"
sips -z 32 32 icon_large.png --out "$ICONSET/icon_16x16@2x.png"
sips -z 32 32 icon_large.png --out "$ICONSET/icon_32x32.png"
sips -z 64 64 icon_large.png --out "$ICONSET/icon_32x32@2x.png"
sips -z 128 128 icon_large.png --out "$ICONSET/icon_128x128.png"
sips -z 256 256 icon_large.png --out "$ICONSET/icon_128x128@2x.png"
sips -z 256 256 icon_large.png --out "$ICONSET/icon_256x256.png"
sips -z 512 512 icon_large.png --out "$ICONSET/icon_256x256@2x.png"
sips -z 512 512 icon_large.png --out "$ICONSET/icon_512x512.png"
sips -z 1024 1024 icon_large.png --out "$ICONSET/icon_512x512@2x.png"

# 使用iconutil将iconset转换为icns
iconutil -c icns "$ICONSET"

# 清理临时文件
rm -rf "$ICONSET"
rm icon_large.png

echo "转换完成: icon.icns 已创建"