# Mission 12458 招商文件

成品：

- `output/Mission_12458_综合产业合作招商书_2026.pptx`

内容定位：

- 面向通用产业、商业、运营及开发合作方；
- 不指向任何特定电商或物流企业；
- 将已确认的CH1开发基础、RU36长期储备和待审批事项明确区分；
- 概念图片不代表现状、许可结果或交付标准。

重新生成：

```bash
python3 -m pip install --target .vendor -r requirements.txt
PYTHONPATH=.vendor python3 scripts/prepare_assets.py
PYTHONPATH=.vendor python3 scripts/build_deck.py
```

`prepare_assets.py`需要本次任务中上传的原始PDF，默认从 Cursor 上传目录读取。
