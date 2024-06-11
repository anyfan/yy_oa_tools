# UTF-8
import re
import subprocess
from string import Template

git_tag = (
    subprocess.check_output(["git", "describe", "--tags", "--abbrev=0"])
    .decode()
    .strip()
)

# git_tag = Major.Minor.Revision[classify]
# v0.0.0α = 0 0 0 α
match = re.match(r"^v(\d+)\.(\d+)\.(\d+)([\D]+)$", git_tag)

version_data = {
    "git_tag": git_tag,
    "Major": int(match.group(1)),  # 主要
    "Minor": int(match.group(2)),  # 次要
    "Revision": int(match.group(3)),  # 修订
    "Classify": ord(match.group(4)),  # 分类
}

with open("template/win_version.py", "rb") as fp:
    text = fp.read().decode("utf-8")

ntext = Template(text).substitute(version_data)

# 保存到文件中
with open("build/win_version.py", "w", encoding="utf-8") as fp:
    fp.write(ntext)