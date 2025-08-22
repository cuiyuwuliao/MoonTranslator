Author: MoonPika

请运行后将鼠标悬停在"操作说明"按钮上, 查看使用说明


-------------------下面是配置说明---------------------
以下所有配置都需要是json格式,配置文件需要以.json为后缀

1. 配置左侧页签
用于配置左侧的页签，每个页签可以包含多个文件按钮。

可配置参数说明(*为必填)
-DisplayText: 进入页签后显示的内容，可以不配置。
-*ContextMenu: 页签的名称，用户在左侧选择的标签。
-*Link: 固定值为 "header"。
-AllowedExtensions: 根目录存在与ContextMenu同名文件夹时, 将扩展名匹配的文件导入为按钮

示例
{
  "BoxItems": [
    {
      "DisplayText": "进入页签后显示的内容，可以不配置",
      "ContextMenu": "左侧页签1",
      "Link": "header",
      "AllowedExtensions": ".mp3, .rpp, .txt, .exe"
    },
    {
      "DisplayText": "进入页签后显示的内容，可以不配置",
      "ContextMenu": "左侧页签2",
      "Link": "header"
    }
  ]
}


2. 配置按钮
用于配置在页签中显示的按钮，拖入文件后或粘贴链接后会自动生成，之后可以手动修改名称或删除不要的配置，配置文件修改后需要重启后生效

可配置参数说明(*为必填)
-DisplayText: 按钮显示的文字。
-*ContextMenu: 按钮所属的页签名称。
-*Link: 需要运行的文件的路径，可以使用相对路径，开头用 ${} 表示，也可以是网址或网盘路径
-IsLarge: 布尔值, 是否使用大按钮，默认为 null。

示例
{
  "BoxItems": [
    {
      "DisplayText": "按钮显示的文字",
      "ContextMenu": "左侧页签1",
      "Link": "header",
      "IsLarge": null
    }
  ]
}


3. 全局设置 masterConfig.json
该配置文件用于全局设置，运行时如果不存在会自动生成。

可配置参数说明(*为必填)
-WorkFolder: 当该应用位于这个文件夹层级下时，拖入该文件夹的文件时，将自动保存为相对路径。
-*EnableAutoRun: 运行时自动启动AutoRun。

示例
{
  "WorkFolder": "binFolder",
  "EnableAutoRun": false
}


4. 配置自动定时无窗运行的程序 AutoRun
该配置文件放在根目录的 AutoRun 文件夹内，如果不存在，请手动创建。

AutoRun弹窗功能：
-可以读取脚本的standard output并弹窗，例如python中的print("一般的内容", flush = true)
-显示限时弹窗: print("!_这里的文字会是弹窗的内容!_", flush = true)
-显示警告弹窗: print("!!_这里的文字会是警告的内容!_", flush = true)

可配置参数说明(*为必填)
-*Link: 需要运行的文件的路径，若使用相对路径，开头用 ${} 表示。
-*AutoRun:
    配置为 -1 时只在启动时运行。
    配置为 0 时在每次 AutoRun 被设置为 ON 时运行。
    配置为大于 0 的数字时为触发间隔（例如 30 代表每 30 秒运行一次，）。

示例
{
  "Link": "${bin}\\Debug\\net6.0-windows7.0\\AutoRun\\earTrainer_demo.exe",
  "AutoRun": 0
}

