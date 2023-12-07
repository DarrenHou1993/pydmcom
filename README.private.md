# 内部的一些说明

## 脚本
```bash
# 构建python包
python setup.py sdist bdist_wheel
# 上传文件
twine upload dist/*

```