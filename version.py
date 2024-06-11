# UTF-8
import re
import subprocess
from string import Template

# 获取当前分支最新tag
git_tag = (
    subprocess.check_output(["git", "describe", "--tags", "--abbrev=0"])
    .decode()
    .strip()
)

# 获取领先tag多少次提交
commits_ahead = (
    subprocess.check_output(["git", "rev-list", "HEAD", "--not", "--tags", "--count"])
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
    "Revision": int(match.group(3))+int(commits_ahead),  # 修订
    "Classify": match.group(4),  # 分类
}
version_data["Classify_num"] = ord(match.group(4))
version_data["version"] = f"v{version_data['Major']}.{version_data['Minor']}.{version_data['Revision']}{version_data['Classify']}"


with open("template/win_version.py", "rb") as fp:
    text = fp.read().decode("utf-8")

ntext = Template(text).substitute(version_data)

# 保存到文件中
with open("build/win_version.py", "w", encoding="utf-8") as fp:
    fp.write(ntext)
